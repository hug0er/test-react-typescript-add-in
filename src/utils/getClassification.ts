import { getDocumentType } from "./getDocumentType";
import JSZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { ClassificationType } from "../types/ClassificationType";
import { wordsInPixel } from "../constants/wordsInPixel";
import { evaluateIncludesInList } from "./utils";
import { wordsInMetada } from "../constants/wordsInMetada";
import { getAttachmentContent } from "./readAttachments";
import * as pdfJS from "pdfjs-dist/legacy/build/pdf";
import PDFJSWorker from "pdfjs-dist/legacy/build/pdf.worker.entry";
pdfJS.GlobalWorkerOptions.workerSrc = PDFJSWorker; // this is to solve a problem with de global worker https://github.com/mozilla/pdf.js/issues/8305

export const getClassification = async (attachment: Office.AttachmentDetailsCompose): Promise<ClassificationType> => {
  try {
    switch (attachment.attachmentType) {
      case Office.MailboxEnums.AttachmentType.Cloud:
        return null;

      case Office.MailboxEnums.AttachmentType.Item:
        return null;

      case Office.MailboxEnums.AttachmentType.File:
        switch (getDocumentType(attachment)) {
          case "pptx":
          case "docx":
            const content = await getAttachmentContent(attachment.id);
            return getOfficeDocClassification(content);
          case "pdf":
            const contentpdf = await getAttachmentContent(attachment.id);
            const pdfClassification = await getPdfClassification(contentpdf);
            return pdfClassification;
          case "xlsx":
            const contentXlsx = await getAttachmentContent(attachment.id);
            const xlsxClassification = await getXslxDocClassification(contentXlsx);
            return xlsxClassification;
          default:
            return null;
        }

      default:
        // Handle attachment formats that are not supported.
        return null;
    }
  } catch (err) {
    console.debug("Error clasificando", err);
    return null;
  }
};

const getOfficeDocClassification = (base64Content: string) => {
  if (!base64Content) return null; //content can be null
  const text = window.atob(base64Content); //base64 to rawText
  const zip = new JSZip(text);
  const doc = new Docxtemplater().loadZip(zip);
  const pixelClassification = getClassificationByText(doc.getFullText());
  if (pixelClassification) return pixelClassification; // This is to evaluate if classification exists can be null for
  const propertyDocument = zip.files["docProps/custom.xml"];
  if (propertyDocument) return getClassificationByText(propertyDocument.asText(), false); // Is not from pixel this is important due to the text in document can contain similar spelling than custom properties
  // can be null
  return null;
};

const getXslxDocClassification = async (base64Content: string) => {
  if (!base64Content) return null; //content can be null
  const xlsx = await import("xlsx");
  const text = window.atob(base64Content); //base64 to rawText
  const doc: any = xlsx.read(base64Content, { type: "base64", bookDeps: true, bookFiles: true });
  const decodedPropsFile = new TextDecoder().decode(doc.files["docProps/custom.xml"].content);
  return getClassificationByText(decodedPropsFile, false);
};

const getPdfClassification = async (base64Content: string) => {
  const text = window.atob(base64Content); //base64 to rawText
  const pdf = await pdfJS.getDocument({ data: text }).promise;

  const pdfData = await (await pdf.getPage(1)).getTextContent();
  const pdfText = pdfData.items.map((item: any) => item.str).join(" "); //any type due to a typescript error with the lib

  const pixelClassification = getClassificationByText(pdfText);
  if (pixelClassification) return pixelClassification;
  const metadata = await pdf.getMetadata();
  const metadataAsString = JSON.stringify(metadata.info); // in case the key or the subthree change in the future
  return getClassificationByText(metadataAsString, false);
};

const getClassificationByText = (str: string, isPixel = true): ClassificationType => {
  if (isPixel) {
    if (evaluateIncludesInList(wordsInPixel.confidential, str)) return "confidential";
    if (evaluateIncludesInList(wordsInPixel.internalUse, str)) return "internalUse";
    if (evaluateIncludesInList(wordsInPixel.restricted, str)) return "restricted";
    return null;
  }
  if (evaluateIncludesInList(wordsInMetada.confidential, str)) return "confidential";
  if (evaluateIncludesInList(wordsInMetada.internalUse, str)) return "internalUse";
  if (evaluateIncludesInList(wordsInMetada.restricted, str)) return "restricted";

  return null;
};
