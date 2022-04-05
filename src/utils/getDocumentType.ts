import { DocumentType } from "../types/DocumentType";

export const getDocumentType = (file: Office.AttachmentDetailsCompose) =>
  file.name.split(".").pop().toLowerCase() as DocumentType;
