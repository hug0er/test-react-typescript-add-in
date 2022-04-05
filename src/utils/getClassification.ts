import { getDocumentType } from "./getDocumentType";
import { Buffer } from "buffer";
import JSZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { ClassificationType } from "../types/ClassificationType";
import { wordsInPixel } from "../constants/wordsInPixel";
import { evaluateIncludesInList } from "./utils";
import { wordsInMetada } from "../constants/wordsInMetada";
import { getAttachmentContent } from "./readAttachments";

export const getClassification = async (attachment: Office.AttachmentDetailsCompose) => {
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
            return null;
          case "xlsx":
            return null;
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
  const propertyClassification = getClassificationByText(propertyDocument.asText(), false); // Is not from pixel this is important due to the text in document can contain similar spelling than custom properties
  return propertyClassification; // can be null
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
