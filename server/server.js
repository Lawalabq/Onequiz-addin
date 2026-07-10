import express from "express";
import cors from "cors";
import "dotenv/config";
import morgan from "morgan";
import { generateQuiz } from "./generateQuiz.js";

const app = express();
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Private-Network", "true");
  next();
});
app.use(cors());
app.use(morgan("dev"));
app.use(express.json());


app.post("/api/generate-quiz", async (req, res) => {
  try {
    const { firstText, secondText, numQuestions } = req.body;

    if (!firstText) {
      return res.status(400).json({ error: "firstText is required" });
    }
    console.log(firstText, secondText, numQuestions)

    const quiz = await generateQuiz(numQuestions, firstText, secondText);
    res.json({ quiz });
    console.log(quiz)
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate quiz" });
  }
});

const port = process.env.PORT || 3001;
app.listen(port, () => console.log(`Server running on port ${port}`));