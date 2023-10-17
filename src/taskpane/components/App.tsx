import * as React from "react";
import { TextField, Stack } from "@fluentui/react";
// import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { Image, IImageProps, ImageFit } from "@fluentui/react/lib/Image";

// These props are defined up here so they can easily be applied to multiple Images.
// Normally specifying them inline would be fine.
const imageProps: IImageProps = {
  imageFit: ImageFit.cover,
  //width: 150,
  height: 100,
  // Show a border around the image (just for demonstration purposes)
  styles: (props) => ({ root: { border: "1px solid " + props.theme.palette.neutralSecondary } }),
};

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  auftrag: number;
  afk: number;
  produktname: string;
  kundenname: string;
}

export interface AppState {
  listItems: HeroListItem[];
  auftrag: "";
  afk: "";
  produktname: "";
  kundenname: "";
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      auftrag: "",
      afk: "",
      produktname: "",
      kundenname: "",
    };
  }
  formData = {
    auftrag: "000",
    afk: "000",
    produktname: "Produktname",
    kundenname: "Kundenname",
  };
  // Handle changes in the input fields
  handleChange = (event) => {
    const target = event.target;
    const value = target.value;
    const name = target.name;
    this.formData[name] = value;

    // this.setState({[name] = value});
  };

  // handleSubmit = async () => {
  //   try {
  //     await Word.run(async (context) => {
  //       await context.sync();
  //     });
  //   } catch (error) {
  //     error(error);
  //   }
  // };

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

  // click = async () => {
  //   return Word.run(async (context) => {
  //     context.document.body.clear();
  //     await context.sync();
  //     const paragraph = context.document.body.insertParagraph(
  //       this.formData.auftrag +
  //         "-A" +
  //         "AFK-" +
  //         this.formData.afk +
  //         "Produkt:" +
  //         this.formData.produktname +
  //         "Kunde:" +
  //         this.formData.kundenname,
  //       Word.InsertLocation.end
  //     );
  //     await context.sync();
  //   });
  // };
  //Palettenboxen mit Kapselsäcken
  clickKapseln = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      /**
       * Insert Word code here
       */
      //Paragraph0
      // insert a paragraph at the start of the document.
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // change the paragraph color to Limnette.
      paragraph.font.color = "#BED200";
      // change the paragraph size to 72.
      paragraph.font.size = 72;
      // change the paragraph font family to Montserrat.
      paragraph.font.name = "Montserrat ExtraBold";
      // change the paragraph text align to center.
      paragraph.alignment = "Centered";

      //Paragraph1
      // insert a empty paragraph at the end of the document. size to 28. and center.
      //const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph1.font.color = "#000000";
      //paragraph1.font.size = 28;
      //paragraph1.font.name = "Montserrat ExtraBold";
      //paragraph1.alignment = "Centered";

      //Paragraph2
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph2 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph2.font.color = "#000000";
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";

      //Paragraph3
      // insert a paragraph at the end of the document. change color to Black. size to 28. font to Montserrat. and center.
      const paragraph3 = context.document.body.insertParagraph("Pulver-Kapseln", Word.InsertLocation.end);
      paragraph3.font.color = "#000000";
      paragraph3.font.size = 28;
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
      const paragraph6 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
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
      const paragraph9 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph9.font.color = "#000000";
      paragraph9.font.size = 72;
      paragraph9.font.name = "Montserrat ExtraBold";
      paragraph9.alignment = "Centered";

      //Paragraph10
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph10 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
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
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
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
      // insert a empty paragraph at the end of the document. size to 28. and center.
      // const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      // paragraph1.font.color = "#000000";
      // paragraph1.font.size = 28;
      // paragraph1.font.name = "Montserrat ExtraBold";
      // paragraph1.alignment = "Centered";

      //Paragraph2
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph2 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph2.font.color = "#000000";
      paragraph2.font.size = 48;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";

      //Paragraph3
      // insert a paragraph at the end of the document. change color to Black. size to 48. font to Montserrat. and center.
      const paragraph3 = context.document.body.insertParagraph("Pulver-Kapseln", Word.InsertLocation.end);
      paragraph3.font.color = "#000000";
      paragraph3.font.size = 28;
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
      const paragraph6 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
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
      const paragraph9 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph9.font.color = "#000000";
      paragraph9.font.size = 72;
      paragraph9.font.name = "Montserrat ExtraBold";
      paragraph9.alignment = "Centered";

      //Paragraph10
      // insert a paragraph at the end of the document. change color to Black. size to 72. font to Montserrat. and center.
      const paragraph10 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
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
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph.font.color = "#BED200";
      paragraph.font.size = 72;
      paragraph.font.name = "Montserrat ExtraBold";
      paragraph.alignment = "Centered";
      //Paragraphen - Dosen ohne Etikett
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.color = "#000000";
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("Pulver-Kapsel-Dosen", Word.InsertLocation.end);
      paragraph1.font.size = 20;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("ohne Etikett", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 36;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Bio - Dosen ohne Etikett
  clickBioDosenohEtt = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph - BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen ohne Etikett
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.color = "#000000";
      paragraph00.font.size = 28;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      const paragraph1 = context.document.body.insertParagraph("Pulver-Kapsel-Dosen", Word.InsertLocation.end);
      paragraph1.font.size = 20;
      paragraph1.font.name = "Montserrat ExtraBold";
      paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("ohne Etikett", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 36;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Dosen Bedruckt
  clickDosenbedruckt = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen Bedruckt
      //const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph00.font.size = 36;
      //paragraph00.font.name = "Montserrat ExtraBold";
      //paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.color = "#000000";
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      //  const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //  paragraph1.font.size = 48;
      //  paragraph1.font.name = "Montserrat ExtraBold";
      //  paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Dosen - Bedruckt", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 36;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Bio - Dosen Bedruckt
  clickBioDosenbedruckt = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph - BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Dosen Bedruckt
      const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph00.font.color = "#000000";
      paragraph00.font.size = 36;
      paragraph00.font.name = "Montserrat ExtraBold";
      paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      //const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph1.font.size = 48;
      //paragraph1.font.name = "Montserrat ExtraBold";
      //paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("Dosen - Bedruckt", Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 36;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph4.font.size = 48;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph6.font.size = 72;
      paragraph6.font.name = "Montserrat ExtraBold";
      paragraph6.alignment = "Centered";
      const paragraph7 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph7.font.size = 72;
      paragraph7.font.name = "Montserrat ExtraBold";
      paragraph7.alignment = "Centered";
      await context.sync();
    });
  };

  //Hochregal Paletten
  clickHochregal = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph Leer - ohne BIO Kennzeichnung
      const paragraphBio = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraphBio.font.color = "#BED200";
      paragraphBio.font.size = 72;
      paragraphBio.font.name = "Montserrat ExtraBold";
      paragraphBio.alignment = "Centered";
      //Paragraphen - Hochregal Palette
      //const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph00.font.color = "#000000";
      //paragraph00.font.size = 36;
      //paragraph00.font.name = "Montserrat ExtraBold";
      //paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.color = "#000000";
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      //const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph1.font.size = 36;
      //paragraph1.font.name = "Montserrat ExtraBold";
      //paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 28;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph4.font.color = "#000000";
      paragraph4.font.size = 72;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph5.font.color = "#000000";
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph6.font.size = 26;
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

  //Hochregal Bio Paletten
  clickBioHochregal = async () => {
    return Word.run(async (context) => {
      context.document.body.clear();
      context.document.body.font.size = 6;
      await context.sync();
      //Paragraph - BIO Kennzeichnung
      const paragraph = context.document.body.insertParagraph("BIO", Word.InsertLocation.end);
      paragraph.font.color = "#BED200";
      paragraph.font.size = 72;
      paragraph.font.name = "Montserrat ExtraBold";
      paragraph.alignment = "Centered";
      //Paragraphen - Hochregal Palette
      //const paragraph00 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph00.font.size = 36;
      //paragraph00.font.name = "Montserrat ExtraBold";
      //paragraph00.alignment = "Centered";
      const paragraph0 = context.document.body.insertParagraph("" + this.formData.produktname, Word.InsertLocation.end);
      paragraph0.font.color = "#000000";
      paragraph0.font.size = 48;
      paragraph0.font.name = "Montserrat ExtraBold";
      paragraph0.alignment = "Centered";
      //const paragraph1 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      //paragraph1.font.size = 36;
      //paragraph1.font.name = "Montserrat ExtraBold";
      //paragraph1.alignment = "Centered";
      const paragraph2 = context.document.body.insertParagraph("" + this.formData.kundenname, Word.InsertLocation.end);
      paragraph2.font.size = 36;
      paragraph2.font.name = "Montserrat ExtraBold";
      paragraph2.alignment = "Centered";
      const paragraph3 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph3.font.size = 28;
      paragraph3.font.name = "Montserrat ExtraBold";
      paragraph3.alignment = "Centered";
      const paragraph4 = context.document.body.insertParagraph("AFK-" + this.formData.afk, Word.InsertLocation.end);
      paragraph4.font.color = "#000000";
      paragraph4.font.size = 72;
      paragraph4.font.name = "Montserrat ExtraBold";
      paragraph4.alignment = "Centered";
      const paragraph5 = context.document.body.insertParagraph(this.formData.auftrag + "-A", Word.InsertLocation.end);
      paragraph5.font.color = "#000000";
      paragraph5.font.size = 36;
      paragraph5.font.name = "Montserrat ExtraBold";
      paragraph5.alignment = "Centered";
      const paragraph6 = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph6.font.size = 28;
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
  // <DefaultButton onClick={this.handleSubmit}>Submit</DefaultButton>

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
          <h3 className="ms-welcome__header">Formular</h3>
          <Stack tokens={{ maxWidth: 300, childrenGap: 10 }}>
            <TextField underlined label="Produkt:" name="produktname" onChange={this.handleChange} />
            <TextField underlined label="Kunde:" name="kundenname" onChange={this.handleChange} />
            <span className="ms-template-list">
              <TextField underlined name="auftrag" onChange={this.handleChange} suffix="-A" />
              <TextField underlined name="afk" onChange={this.handleChange} prefix="AFK-" />
            </span>
          </Stack>
          <p className="ms-welcome__anleitung ms-font-s">
            <b>Klicke</b> auf eine Vorlage um sie zu <b>laden</b>.
          </p>
          <br />
          <h2>Produktion</h2>
          <h3>Palettenboxen</h3>
          <span className="ms-template-list">
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickKapseln}
              {...imageProps}
              src="./../../../assets/kapselscke.png"
              alt="Palettenbox mit Kapselscken"
            ></Image>
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickBioKapseln}
              {...imageProps}
              src="./../../../assets/kapselsckebio.png"
              alt="Palettenbox mit Kapselscken - Bio"
            ></Image>
          </span>
          <h3>Kartons</h3>
          <h4>Kapseldosen ohne Etikett</h4>
          <span className="ms-template-list">
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickDosenohEtt}
              {...imageProps}
              src="./../../../assets/KapseldosenohneEtikett.png"
              alt="Kapseldosen ohne Etikett"
            ></Image>
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickBioDosenohEtt}
              {...imageProps}
              src="./../../../assets/KapseldosenohneEtikettBio.png"
              alt="Kapseldosen ohne Etikett - Bio"
            ></Image>
          </span>
          <h4>Dosen bedruckt</h4>
          <span className="ms-template-list">
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickDosenbedruckt}
              {...imageProps}
              src="./../../../assets/Dosenbedruckt.png"
              alt="Dosen bedruckt"
            ></Image>
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickBioDosenbedruckt}
              {...imageProps}
              src="./../../../assets/DosenbedrucktBio.png"
              alt="Dosen bedruckt - Bio"
            ></Image>
          </span>
          <h2>Lager</h2>
          <h3>Hochregal - Palette</h3>
          <span className="ms-template-list">
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickHochregal}
              {...imageProps}
              src="./../../../assets/HLVerdoselteWare.png"
              alt="Verdoselte Ware"
            ></Image>
            <Image
              className="ms-welcome__imageaction"
              onClick={this.clickBioHochregal}
              {...imageProps}
              src="./../../../assets/HLVerdoselteWareBio.png"
              alt="Verdoselte Bio Ware"
            ></Image>
          </span>
        </HeroList>
      </div>
    );
  }
}
