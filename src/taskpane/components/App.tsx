import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global Word, require */

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

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Kapselsäcke Konventionell", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  click1 = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Kapselsäcke - Bio", Word.InsertLocation.end);

      // change the paragraph color to green.
      paragraph.font.color = "green";

      await context.sync();
    });
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
        <HeroList
          message="Unternehmensweite Standardvorlagen zur Vereinheitlichung und Gleichschaltung"
          items={this.state.listItems}
        >
          <p className=".ms-welcome__anleitung ms-font-s">
            Durch <b>anklicken</b> einer Vorlage wird diese geladen und das geöffnete Dokument <b>überschrieben</b>.
          </p>
          <h2>Produktion - Palettenboxen</h2>
          <span className="ms-template-list">
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapselsäcke
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.click1}
            >
              Kapselsäcke - Bio
            </DefaultButton>
          </span>
          <h2>Produktion - Kartons</h2>
          <span className="ms-template-list">
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen ohne Etikett
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.click1}
            >
              Kapseldosen ohne Etikett - Bio
            </DefaultButton>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Dosen bedruckt
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.click1}
            >
              Dosen bedruckt - Bio
            </DefaultButton>
          </span>
          <h2>Lager - Hochregal</h2>
          <span className="ms-template-list">
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.click1}
            >
              Kapseldosen - Bio
            </DefaultButton>
          </span>
        </HeroList>
      </div>
    );
  }
}
