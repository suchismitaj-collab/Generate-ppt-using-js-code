const pptxgen = require("pptxgenjs");

async function createPresentation() {
    let pptx = new pptxgen();

    pptx.layout = "LAYOUT_WIDE";

    let slide = pptx.addSlide();

    slide.addText("Types of Arduino", {
        x: 1,
        y: 1,
        w: 5,
        h: 1,
        fontSize: 24,
        bold: true
    });

    slide.addText("Major arduino types include: ", {
        x: 1,
        y: 2,
        w: 6,
        h: 1,
    });

    await pptx.writeFile({ fileName: "Arduino_presentation.pptx" });
    console.log("File saved successfully!");
}

createPresentation().catch(console.error);