import { OpenAI } from "openai";
import fs from "fs";




const prompt = fs.readFileSync(new URL("./prompt.txt", import.meta.url), "utf8");

const client = new OpenAI({
	baseURL: "https://router.huggingface.co/v1",
	apiKey: process.env.HF_TOKEN,
});




export async function generateQuiz(numQuestions = 10, firstText, secondText = "") {
  const completePrompt = prompt
    .replace("**{firstText}**", firstText)
    .replace("**{secondText}**", secondText)
    .replace("**{NUMBER_OF_QUESTIONS}**", numQuestions);

    const chatCompletion = await client.chat.completions.create({
    model: "openai/gpt-oss-120b:groq",
      messages: [
          {
              role: "user",
              content: completePrompt,
          },
      ],
});

  return chatCompletion.choices[0].message.content;
}
















