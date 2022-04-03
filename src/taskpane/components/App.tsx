import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { useState } from "react";
import Profile from "./Profile";
import Read from "../read";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

const App: React.FC<AppProps> = (props) => {
  const { title, isOfficeInitialized } = props;
  const [listItems, setListItems] = useState([
    {
      icon: "Ribbon",
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: "Unlock",
      primaryText: "Unlock features and functionality",
    },
    {
      icon: "Design",
      primaryText: "Create and visualize like a pro",
    },
  ]);

  const click = async () => {
    /**
     * Insert your Outlook code here
     */
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={props.title} message="Welcome" />
      <Profile />
      <Read />
      {/* <p className="ms-font-l">
        Ejemplo leer correos <b>Run</b>.
      </p>
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        Run
      </DefaultButton> */}
    </div>
  );
};
export default App;
