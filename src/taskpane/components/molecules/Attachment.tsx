import { ImageIcon } from "@fluentui/react/lib/Icon";
import React, { useEffect, useState } from "react";
import { getClassification } from "../../../utils/getClassification";
import { getDocumentType } from "../../../utils/getDocumentType";
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { Stack, IStackStyles, IStackTokens } from '@fluentui/react/lib/Stack';
import { Text, ITextProps } from '@fluentui/react/lib/Text';

type AttachmentProps = {
  attachment: Office.AttachmentDetailsCompose;
};

const classNames = mergeStyleSets({
  one: {
    width: 48,
    height: 44,
    marginLeft: 0,
  },
});

const token: IStackTokens = {
  childrenGap: 5,
  padding: 10,
};

const stackStyles: IStackStyles = {
  root: {
    borderBottom: '1px solid rgba(10, 10, 10, 0.2)',
  },
};

const Attachment: React.FC<AttachmentProps> = ({ attachment }) => {
  const { id, name} = attachment;
  const [classificaton, setClassificaton] = useState(null);
  const [iconUrl, setIconUrl] = useState(null);
  const [docType, setDocType] = useState(null);

  const getFileIcon = (docType: string) => `https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/20/${docType}.svg`;


  useEffect(() => {
    getClassification(attachment).then((classificaton) => {
      setClassificaton(classificaton || "no existe")
    });
    const typeFile = getDocumentType(attachment)
    setDocType(typeFile);
    setIconUrl(getFileIcon(typeFile));
  }, [id]);

  return (
    <Stack horizontal tokens={token} styles={stackStyles}>
      <Stack>
        <ImageIcon
          className={classNames.one}
          imageProps={{
            src: iconUrl,
            alt: `${docType} file icon`,
          }}
        />
      </Stack>
      <Stack>
        <Text key={name + id} variant={'mediumPlus' as ITextProps['variant']} block>
          <b>{name}</b>
        </Text>
        <Text key={classificaton + id} variant={'medium' as ITextProps['variant']} block>
          Categoria:
          {classificaton == 'internalUse' && 
          <span style={{color: 'green', fontWeight: 650}}> {`${classificaton}`}</span>}
          
          {classificaton != 'internalUse' && 
          <span style={{color: 'red', fontWeight: 650}}> {`${classificaton}`}</span>}
        </Text>
      </Stack>
    </Stack>
  );
};

export default Attachment;
