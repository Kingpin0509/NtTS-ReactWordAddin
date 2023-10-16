/* eslint-disable no-undef */
import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as excel from "./Excel.App";
import * as onenote from "./OneNote.App";
import * as outlook from "./Outlook.App";
import * as powerpoint from "./PowerPoint.App";
import * as project from "./Project.App";
import * as word from "./App";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Steigerung der Betriebsqualität",
        },
        {
          icon: "Unlock",
          primaryText: "Freigeben von Unternehmensressourcen",
        },
        {
          icon: "Design",
          primaryText: "Förderung einheitlicher Firmenkultur und -identität",
        },
      ],
    });
  }

  //onClick Platzhalter
  click = async () => {
    switch (Office.context.host) {
      case Office.HostType.Excel: {
        const excelApp = new excel.default(this.props, this.context);
        return excelApp.click();
      }
      case Office.HostType.OneNote: {
        const onenoteApp = new onenote.default(this.props, this.context);
        return onenoteApp.click();
      }
      case Office.HostType.Outlook: {
        const outlookApp = new outlook.default(this.props, this.context);
        return outlookApp.click();
      }
      case Office.HostType.PowerPoint: {
        const powerpointApp = new powerpoint.default(this.props, this.context);
        return powerpointApp.click();
      }
      case Office.HostType.Project: {
        const projectApp = new project.default(this.props, this.context);
        return projectApp.click();
      }
      case Office.HostType.Word: {
        const wordApp = new word.default(this.props, this.context);
        return wordApp.click();
      }
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Initialisierung Fehlgeschlagen - Body leer - sideload?"
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/logo-filled.png")}
          title={this.props.title}
          message="Herzlich Willkommen"
        />
        <HeroList message="One-Klick A4 Standardvorlagen" items={this.state.listItems}>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Test
          </DefaultButton>
          <p className="ms-welcome__anleitung ms-font-s">
            <b>Klicke</b> auf eine Vorlage um sie zu <b>laden</b>.
          </p>
          <br />
        </HeroList>
      </div>
    );
  }
}
