import React from "react";
import Attachment from "./Attachment";

const Attachments: React.FC = () => {
  const attachments = Office.context.mailbox.item.attachments;
  Office.context.mailbox.userProfile;
  console.log("componente attachments", document);
  return (
    <>
      {attachments.map((attachment) => (
        <Attachment attachment={attachment} key={attachment.id} />
      ))}
    </>
  );
};

export default Attachments;
