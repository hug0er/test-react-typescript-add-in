import React, { useEffect, useState } from "react";
import { getAttachments } from "../../utils/readAttachments";
import Attachment from "./Attachment";

const Attachments: React.FC = () => {
  const [attachments, setAttachments] = useState<Office.AttachmentDetailsCompose[]>([]);
  useEffect(() => {
    getAttachments().then((list) => setAttachments(list));
  }, []);
  Office.context.mailbox.userProfile;
  return (
    <>
      {attachments.map((attachment) => (
        <Attachment attachment={attachment} key={attachment.id} />
      ))}
    </>
  );
};

export default Attachments;
