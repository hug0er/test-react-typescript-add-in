import { isReadMode } from "./currentMode";
import { toPromise } from "./utils";

export const getAttachments = async (): Promise<Office.AttachmentDetailsCompose[]> => {
  if (isReadMode()) return Office.context.mailbox.item.attachments;
  const attachmentsResponse: Office.AsyncResult<Office.AttachmentDetailsCompose[]> = await toPromise(
    Office.context.mailbox.item.getAttachmentsAsync,
    {}
  );
  const attachments = attachmentsResponse.value;
  return attachments;
};

export const getAttachmentContent = async (id: string) => {
  const content = await toPromise(Office.context.mailbox.item.getAttachmentContentAsync, id);

  switch (content.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      return content.value.content;

    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      return null;

    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      return null;

    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      return null;
    default:
      // Handle attachment formats that are not supported.
      return null;
  }
}; // as base64
