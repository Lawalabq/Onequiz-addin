export async function getImageOcrData(context,page){
/* This function parses throught the active page and returns the text ocr recognised from oneNote
I need to add a check that if it doesnt exist it shoudl reuest an external API to do OCR 
*/



  const outlines = page.contents; // PageContentCollection
  outlines.load("items");
  await context.sync();


  //Store image OCR
  const images = [];

  for (const content of outlines.items) {
    if (content.type === "Image") {
    // store if type imge
      const img = content.image;
      img.load("id,width,height,ocrData");
      images.push(img);
    } else if (content.type === "Outline") {
        //check if outline has images
      const outline = content.outline;
      outline.paragraphs.load("items");
      await context.sync();
      for (const p of outline.paragraphs.items) {
        if (p.type === "Image") {
          const img = p.image;
          img.load("id,width,height,ocrData");
          images.push(img);
        }
      }
    }
  }

  await context.sync();

  const results = images.map(img => ({
    ocrText: img.ocrData ? img.ocrData.ocrText : null,
  }));

  // Base64 needs a separate call per image since getBase64Image() returns a ClientResult
  const base64Results = images.map(img => img.getBase64Image());
  await context.sync();


  // #Add check for null then use my function

  // for (var i =0 ; i<results.length;i++){
  //   if (results[i].ocrText == null){

  //     // const res = await fetch("api/base64-OCR", {
  //   //   method: "POST",
  //   //   headers: { "Content-Type": "application/json" },
  //   //   body: JSON.stringify({ base64Strings: [base64Strings]})
  //   // });

  //     // let json = await res.json();
  //     //  results[i].ocrText = json;

  //   }




  // }

 // return a list of text
  return results.map(img=> img.ocrText);


}


export async function getInkAnalysisResults(context, page, wordAlternatesChoice = 0) {

    // COntext is required for all oneNote interations
    //page is current page
    // wordAlternatesChoice chose how many word probs i need

    // load ink analysis
  const inkAnalysis = page.inkAnalysisOrNull;
  await context.sync();

  if (inkAnalysis.isNull) {
    console.log("No ink analysis available.");
    return "";
  }

  // Load inkstokes
  const paragraphs = inkAnalysis.paragraphs;
  paragraphs.load("items");
  await context.sync();

  for (const paragraph of paragraphs.items) {
    paragraph.lines.load("items");
  }
  await context.sync();

  for (const paragraph of paragraphs.items) {
    for (const line of paragraph.lines.items) {
      line.words.load("items");
    }
  }
  await context.sync();


  // load words
  for (const paragraph of paragraphs.items) {
    for (const line of paragraph.lines.items) {
      for (const word of line.words.items) {
        word.load("wordAlternates");
      }
    }
  }
  await context.sync();

  let recognizedText = "";
  for (const paragraph of paragraphs.items) {
    for (const line of paragraph.lines.items) {
      for (const word of line.words.items) {
        recognizedText += word.wordAlternates[wordAlternatesChoice] + " ";
      }
      recognizedText += "\n";
    }
  }

  return recognizedText;
}









