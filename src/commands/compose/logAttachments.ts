import { Buffer } from "buffer";
import JSZip from "pizzip";
import Docxtemplater from "docxtemplater";

export const logAttachments = (event: Office.AddinCommands.Event) => {
  Office.context.mailbox.item.getAttachmentsAsync((attachments) => {
    return attachments.value.forEach((attachment) => {
      Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (content) => {
        const text = window.atob(content.value.content); //base64 to rawText
        const buffer = Buffer.from(text);
        const blob = new Blob([text]);
        const reader = new FileReader();
        reader.onload = (eventRead) => {
          const content = eventRead.target.result;
          // const metadata = exif.readFromBinaryFile(buffer.buffer);
          // const metadata2 = new Jdataview(buffer);
          const zip = new JSZip(content);
          const doc = new Docxtemplater().loadZip(zip);
          // const text = doc.getFullText();
          // const metadata3 = doc.getZip();

          // const zip2 = new AdmZip(buffer);
          // yauzl.fromBuffer(buffer, (err, file) => {
          //   console.log(err, file);
          // });
          // console.log(text);
          console.log("keys", Object.keys(zip.files).length, zip.comment, JSON.stringify(zip.files));
          let cnt = 1;
          for (const i in zip.files) {
            const string = new TextDecoder().decode(zip.files["docProps/custom.xml"].asUint8Array());
            const text2 = doc.getFullText(i);
            const rawResult = JSON.stringify(zip.files["docProps/custom.xml"]);
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
            console.log("_______________________________________");
            console.log("_______________________________________");
            console.log("_______________________________________");
            console.log(i);
            console.log("_______________________________________");
            console.log(JSON.stringify(string));
            console.log("_______________________________________");
            console.log(text2 ? text2 : "no hay textoooo");

            // console.log("archivo cargado", i, text2, rawResult);
          }
          console.log("hola mundo");
          event.completed();
        };
        reader.readAsText(blob);
      });
    });

    // Be sure to indicate when the add-in command function is complete
  });
};
const write = (message: string) => {
  document.getElementById("message").innerText += message;
};

var JsonToArray = function (json) {
  var str = JSON.stringify(json, null, 0);
  var ret = new Uint8Array(str.length);
  for (var i = 0; i < str.length; i++) {
    ret[i] = str.charCodeAt(i);
  }
  return ret;
};

var binArrayToJson = function (binArray) {
  var str = "";
  for (var i = 0; i < binArray.length; i++) {
    str += String.fromCharCode(parseInt(binArray[i]));
  }
  return JSON.parse(str);
};
