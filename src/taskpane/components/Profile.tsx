import React from "react";
import { DocumentCard } from "@fluentui/react";

const Profile: React.FC = () => {
  const profile = Office.context.mailbox.userProfile;
  return <DocumentCard>Hola, {`${profile.displayName} (${profile.emailAddress})`}</DocumentCard>;
};

export default Profile;
