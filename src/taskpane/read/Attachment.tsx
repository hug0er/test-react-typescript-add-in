import React from "react";
import { Buffer } from "buffer";
import JSZip from "pizzip";
import Docxtemplater from "docxtemplater";

type AttachmentProps = {
  attachment: Office.AttachmentDetails;
};

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  console.log("result");
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
    // Handle attachment formats that are not supported.
  }
}

const Attachment: React.FC<AttachmentProps> = ({
  attachment: { size, id, name, attachmentType, contentType, isInline },
}) => {
  console.log("montando attachment");
  Office.context.mailbox.item.getAttachmentContentAsync(id, (result) => {
    const text = window.atob(result.value.content);
    const buffer = Buffer.from(text);

    const blob = new Blob([text]);
    const reader = new FileReader();
    reader.onload = (eventRead) => {
      const content = eventRead.target.result;
      const zip = new JSZip(content);
      const doc = new Docxtemplater().loadZip(zip);
      const text = doc.getFullText();

      // const zip2 = new AdmZip(buffer);
      // yauzl.fromBuffer(buffer, (err, file) => {
      //   console.log(err, file);
      // });
      // console.log(text);
      console.log("keys", Object.keys(zip.files).length, zip.comment);
      let cnt = 1;
      for (const i in zip.files) {
        const uint8array = new TextEncoder().encode("¢");
        const string = new TextDecoder().decode(zip.files[i].asUint8Array());
        const text2 = doc.getFullText(i);
        const rawResult = JSON.stringify(zip.files[i]);
        const zipText = zip.files[i].asText();
        if (
          string.includes("2-Restringido") ||
          string.includes("3-Confidencial") ||
          text2.includes("2-Restringido") ||
          text2.includes("3-Confidencial") ||
          rawResult.includes("2-Restringido") ||
          rawResult.includes("3-Confidencial") ||
          zipText.includes("2-Restringido") ||
          zipText.includes("3-Confidencial") ||
          text2.includes("KriptosClassAi") ||
          text2.includes("KriptosClassAi") ||
          rawResult.includes("KriptosClassAi") ||
          rawResult.includes("KriptosClassAi") ||
          zipText.includes("KriptosClassAi") ||
          zipText.includes("KriptosClassAi") ||
          string.includes("KriptosClassAi")
        ) {
          // console.log("archivo encontrado en ", i);
        }
        console.log("entonctrado en", i, cnt);
        cnt = cnt + 1;
        // console.log("_______________________________________");
        // console.log("_______________________________________");
        // console.log("_______________________________________");
        // console.log(i);
        // console.log("_______________________________________");
        // console.log(JSON.stringify(string));
        // console.log("_______________________________________");
        // console.log(text2 ? text2 : "no hay textoooo");

        // console.log("archivo cargado", i, text2, rawResult);
      }
      console.log("hola mundo");
    };
    reader.readAsText(blob);

    // getDocumentProperties.fromBuffer(buffer, function (err, data) {
    //   if (err) throw err;
    //   console.log(data);
    //   /*
    //   {
    //     application: 'Microsoft Macintosh PowerPoint',
    //     applicationVersion: '15.0027',
    //     comments: 'This is an example document for testing the retrieval of document properties.',
    //     company: 'Acme',
    //     created: '2017-10-13T15:54:58Z',
    //     createdBy: 'Microsoft Office User',
    //     keywords: 'Example Sample Test Properties',
    //     manager: 'Jane Doe',
    //     modified: '2017-10-13T19:27:15Z',
    //     modifiedBy: 'Microsoft Office User',
    //     notes: 0,
    //     paragraphs: 2,
    //     revision: 3,
    //     slides: 1,
    //     subject: 'Example',
    //     title: 'Test PowerPoint Document',
    //     totalTime: 2,
    //     words: 8
    //   }
    //   */
    // });
  });
  //   const item = Office.context.mailbox.item;
  //   const options = { asyncContext: { currentItem: item } };
  //   item.getAttachmentsAsync(options, (result) => {
  //     if (result.value.length > 0) {
  //       for (let i = 0; i < result.value.length; i++) {
  //         result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id);
  //       }
  //     }
  //   });
  return (
    <div>
      <div>{`Nombre: ${name}`}</div>
      <div>{`ID: ${id}`}</div>
      <div>{`Tamano del archivo: ${size} bytes`}</div>
      <div>{`Tipo del archivo: ${attachmentType}`}</div>
      <div>{`Tipo del contenido: ${contentType}`}</div>
      <div>{`Is inline (si debe estar en el cuerpo o contenido): ${isInline}`}</div>
    </div>
  );
};

export default Attachment;
