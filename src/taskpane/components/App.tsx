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

  clickBioKapseln = async () => {
    return Word.run(async (context) => {
      /**
       *  Word code here
       */

      //Paragraph0
      // insert a paragraph at the start of the document.
      const paragraph = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      // change the paragraph color to Limnette.
      paragraph.font.color = "#BED200";
      // change the paragraph size to 72.
      paragraph.font.size = 72;
      // change the paragraph font family to Montserrat.
      paragraph.font.name = "Montserrat ExtraBold";
      // change the paragraph text align to center.
      paragraph.alignment = "Centered";

      //Paragraph1
      // insert a empty paragraph at the end of the document. size to 48. and center.
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";

      //Paragraph2
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph2 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph2.font.color = "#000000";
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";

      //Paragraph3
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph3 = context.document.body.insertParagraph("Pulver-Kapseln", Word.InsertLocation.end);
      paragraph3.font.color = "#000000";
      paragraph3.font.size = 48;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";

      //Paragraph4
      // insert a empty paragraph at the end of the document. size to 36. and center.
      const paragraph4 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph4.font.size = 36;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";

      //Paragraph5
      // insert a empty paragraph at the end of the document. size to 36.
      // const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // paragraph5.font.size = 36;

      //Paragraph6
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph6 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph6.font.color = "#000000";
      paragraph6.font.size = 48;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";

      //Paragraph7
      // insert a empty paragraph at the end of the document. size to 48. and center.
      const paragraph7 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph7.font.size = 36;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";

      //Paragraph8
      // insert a empty paragraph at the end of the document. size to 48.
      // const paragraph8 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // paragraph8.font.size = 48;

      //Paragraph9
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph9 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph9.font.color = "#000000";
      paragraph9.font.size = 72;
      paragraph9.font.name = "Montserrat ExtraBold";
      paragraph9.alignment = "Centered";

      //Paragraph10
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph10 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph10.font.color = "#000000";
      paragraph10.font.size = 72;
      paragraph10.font.name = "Montserrat ExtraBold";
      paragraph10.alignment = "Centered";

      await context.sync();
    });
  };

  clickKapseln = async () => {
    return Word.run(async (context) => {
      /**
       * Insert Word code here
       */

      //Paragraph0
      // insert a paragraph at the start of the document.
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // change the paragraph color to Limnette.
      // paragraph.font.color = "#BED200";
      // change the paragraph size to 72.
      paragraph.font.size = 72;
      // change the paragraph font family to Montserrat.
      paragraph.font.name = "Montserrat ExtraBold";
      // change the paragraph text align to center.
      paragraph.alignment = "Centered";

      //Paragraph1
      // insert a empty paragraph at the end of the document. size to 48. and center.
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";

      //Paragraph2
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph2 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph2.font.color = "#000000";
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";

      //Paragraph3
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph3 = context.document.body.insertParagraph("Pulver-Kapseln", Word.InsertLocation.end);
      paragraph3.font.color = "#000000";
      paragraph3.font.size = 48;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";

      //Paragraph4
      // insert a empty paragraph at the end of the document. size to 36. and center.
      const paragraph4 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph4.font.size = 36;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";

      //Paragraph5
      // insert a empty paragraph at the end of the document. size to 36.
      // const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // paragraph5.font.size = 36;

      //Paragraph6
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph6 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph6.font.color = "#000000";
      paragraph6.font.size = 48;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";

      //Paragraph7
      // insert a empty paragraph at the end of the document. size to 48. and center.
      const paragraph7 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph7.font.size = 36;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";

      //Paragraph8
      // insert a empty paragraph at the end of the document. size to 48.
      // const paragraph8 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // paragraph8.font.size = 48;

      //Paragraph9
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph9 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph9.font.color = "#000000";
      paragraph9.font.size = 72;
      paragraph9.font.name = "Montserrat ExtraBold";
      paragraph9.alignment = "Centered";

      //Paragraph10
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph10 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph10.font.color = "#000000";
      paragraph10.font.size = 72;
      paragraph10.font.name = "Montserrat ExtraBold";
      paragraph10.alignment = "Centered";
      await context.sync();
    });
  };
  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */
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
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickKapseln}
            >
              Kapselsäcke
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickBioKapseln}
            >
              Kapselsäcke - Bio
            </DefaultButton>
          </span>
          <h2>Produktion - Kartons</h2>
          <span className="ms-template-list">
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen ohne Etikett
            </DefaultButton>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen ohne Etikett - Bio
            </DefaultButton>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Dosen bedruckt
            </DefaultButton>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Dosen bedruckt - Bio
            </DefaultButton>
          </span>
          <h2>Lager - Hochregal</h2>
          <span className="ms-template-list">
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen
            </DefaultButton>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Kapseldosen - Bio
            </DefaultButton>
          </span>
        </HeroList>
      </div>
    );
  }
}
