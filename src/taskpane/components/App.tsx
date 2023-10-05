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

  //onClick Platzhalter
  click = async () => {
    return Word.run(async (context) => {
      await context.sync();
    });
  };

  //Palettenboxen mit Kapselsäcken
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

  //Palettenboxen mit BIO Kapselsäcken
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

  //Dosen ohne Etikett
  clickDosenohEtt = async () => {
    return Word.run(async (context) => {
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph.font.color = "#BED200";
      paragraph.font.size = 72;
      paragraph.font.name = "Montserrat ExtraBold";
      paragraph.alignment = "Centered";
      //Paragraphen - Dosen ohne Etikett
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("Pulver-Kapsel-Dosen", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("ohne Etikett", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 72;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Bio - Dosen ohne Etikett
  clickBioDosenohEtt = async () => {
    return Word.run(async (context) => {
      //Paragraph - BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen ohne Etikett
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("Pulver-Kapsel-Dosen", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("ohne Etikett", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 72;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Bio - Dosen Bedruckt
  clickBioDosenbedruckt = async () => {
    return Word.run(async (context) => {
      //Paragraph - BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen Bedruckt
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Dosen - Bedruckt", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 72;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Dosen Bedruckt
  clickDosenbedruckt = async () => {
    return Word.run(async (context) => {
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen Bedruckt
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 48;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Dosen - Bedruckt", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 72;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Hochregal Bio Paletten
  clickBioHochregal = async () => {
    return Word.run(async (context) => {
      //Paragraph - BIO Kennzeichnung
      const paragraph = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraph.font.color = "#BED200";
      paragraph.font.size = 72;
      paragraph.font.name = "Montserrat ExtraBold";
      paragraph.alignment = "Centered";
      //Paragraphen - Hochregal Palette
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 36;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph4.font.color = "#000000";
      paragraph4.font.size = 72;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph5.font.color = "#000000";
      paragraph5.font.size = 48;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph6.font.size = 36;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("0000 Dosen", Word.InsertLocation.end);
      paragraph7.font.color = "#000000";
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Hochregal Paletten
  clickHochregal = async () => {
    return Word.run(async (context) => {
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Hochregal Palette
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("Produktname", Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph1.font.size = 36;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Kundenname", Word.InsertLocation.end);
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 72;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("AFK-000", Word.InsertLocation.end);
      paragraph4.font.color = "#000000";
      paragraph4.font.size = 72;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("000-A", Word.InsertLocation.end);
      paragraph5.font.color = "#000000";
      paragraph5.font.size = 48;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph6.font.size = 36;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph("0000 Dosen", Word.InsertLocation.end);
      paragraph7.font.color = "#000000";
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
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
            Durch <b>anklicken</b> einer Vorlage wird diese in das geöffnete Dokument <b>eingefügt</b>.
          </p>
          <br></br>
          <h2>Produktion</h2>
          <h3>Palettenboxen mit:</h3>
          <span className="ms-template-list">
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickKapseln}
            >
              Kapselsäcken
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickBioKapseln}
            >
              Kapselsäcken - Bio
            </DefaultButton>
          </span>
          <h2>Produktionsvorbereitung</h2>
          <h3>Kartons mit:</h3>
          <span className="ms-template-list">
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickDosenohEtt}
            >
              Kapseldosen ohne Etikett
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickBioDosenohEtt}
            >
              Kapseldosen ohne Etikett - Bio
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickDosenbedruckt}
            >
              Dosen bedruckt
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickBioDosenbedruckt}
            >
              Dosen bedruckt - Bio
            </DefaultButton>
          </span>
          <h2>Produktionsabschluss</h2>
          <h3>Hochregalpallette mit:</h3>
          <span className="ms-template-list">
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickHochregal}
            >
              Dosenprodukte
            </DefaultButton>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.clickBioHochregal}
            >
              Bio Dosenprodukte
            </DefaultButton>
          </span>
        </HeroList>
      </div>
    );
  }
}
