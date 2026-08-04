import { OpenAI } from "openai";
import fs from "fs";




const prompt = fs.readFileSync(new URL("./prompt.txt", import.meta.url), "utf8");
const quizFromSectionPrompt = fs.readFileSync(new URL ("./section prompt.txt", import.meta.url),'utf8');

const client = new OpenAI({
	baseURL: "https://router.huggingface.co/v1",
	apiKey: process.env.HF_TOKEN,
});


const groq_client  = new OpenAI({
  apiKey:process.env.GROQ_API_KEY,
    baseURL:"https://api.groq.com/openai/v1",

})







//make class??
export async function generateQuiz(numQuestions = 10, firstText, secondText = "", imageOCR="") {
  
    let fullImageOCR ="";
    if (imageOCR!==""){
      for (var i= 0 ; i,imageOCR.length;i++){
        fullImageOCR+=imageOCR[i]
      }
    };
    console.log(fullImageOCR);

    const completePrompt = prompt
    .replace("**{firstText}**", firstText)
    .replace("**{secondText}**", secondText)
    .replace("**{NUMBER_OF_QUESTIONS}**", numQuestions)
    .replace("**{imageOcrText}**",fullImageOCR);


    const chatCompletion = await groq_client.chat.completions.create({
    // model: "openai/gpt-oss-120b:groq",
      model: "openai/gpt-oss-20b", // for groq api

      messages: [
          {
              role: "user",
              content: completePrompt,
          },
      ],
});

  return chatCompletion.choices[0].message.content;
  // return chatCompletion.output_text
}

export async function generateQuizFromSection(numQuestions = 20, inkCharacterText, OCRText) {

  let sectionPrompts ="";

  for(var i =0; i<inkCharacterText.length;i++){
    let sectionInput = `Section ${i+1}:`;
    sectionInput+="inkCharacter: \n";
    sectionInput+=inkCharacterText[i];
    sectionInput+="OCRText: \n";
    sectionInput+=OCRText[i];

    sectionPrompts+=sectionInput;
    sectionPrompts+="\n";
  }

  const completePrompt = quizFromSectionPrompt.replace("**{sectionText}**",sectionPrompts)
  .replace("**{NUMBER_OF_QUESTIONS}**",numQuestions);

  const chatCompletion = await groq_client.chat.completions.create({
    // model: "openai/gpt-oss-120b:groq",
      model: "openai/gpt-oss-20b", // for groq api
      messages: [
          {
              role: "user",
              content: completePrompt,
          },
      ],
});

  return chatCompletion.choices[0].message.content;
  // return chatCompletion.output_text





}
  






export async function getBase64Text(base64Image) {
  const apiKey = process.env.FREE_OCR_API_KEY;
  const url = "https://api.ocr.space/parse/image";
  let response = await fetch(url, {
  method: "POST",
  headers: {
    "apikey": apiKey,
    "Content-Type": "image/png"
  },
  body: JSON.stringify({
    base64Image: f`data:image/png;base64,${base64Image}`
  })
});


  if  (response.ok) {
    const data = await response.json();
    return data;
  }
  else {
    console.error("Error fetching OCR data:", response.status, response.statusText);
    return null;
  }

}














// Loading images
// let image = null;

// await OneNote.run(async (context) => {
//     // Get the current outline.
//     const outline = context.application.getActiveOutline();

//     // Queue a command to load paragraphs and their types.
//     outline.load("paragraphs")
//     await context.sync();

//     for (let i=0; i < outline.paragraphs.items.length; i++) {
//         const paragraph = outline.paragraphs.items[i];
//         if (paragraph.type == "Image")
//         {
//             image = paragraph.image;
//         }
//     }
//     if (image != null) {
//         image.load("ocrData");
//     }

//     await context.sync();
            
//     // Log ocrText and ocrLanguageId.
//     console.log(image.ocrData.ocrText);
//     console.log(image.ocrData.ocrLanguageId);
// });
