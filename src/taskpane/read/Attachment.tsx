import React, { useEffect, useState } from "react";
import { getClassification } from "../../utils/getClassification";
import { getDocumentType } from "../../utils/getDocumentType";

type AttachmentProps = {
  attachment: Office.AttachmentDetailsCompose;
};

const Attachment: React.FC<AttachmentProps> = ({ attachment }) => {
  const { size, id, name, attachmentType, isInline } = attachment;
  const [classificaton, setClassificaton] = useState(null);

  useEffect(() => {
    getClassification(attachment).then((classificaton) => setClassificaton(classificaton || "no existe"));
  }, [id]);

  return (
    <div>
      <div>
        <strong>{`Nombre: ${name}`}</strong>
      </div>
      <div>{`ID: ${id}`}</div>
      <div>{`Tamano del archivo: ${size} bytes`}</div>
      <div>{`Tipo del archivo: ${attachmentType}`}</div>
      <div>{`Tipo del contenido: ${getDocumentType(attachment)}`}</div>
      <div>{`Is inline (si debe estar en el cuerpo o contenido): ${isInline}`}</div>
      <div>{`Clasificacion: ${classificaton}`}</div>
    </div>
  );
};

export default Attachment;
