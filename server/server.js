import express from "express";
import cors from "cors";
import "dotenv/config";
import morgan from "morgan";
import { rateLimit, HOUR } from 'express-rate-limit'
import { generateQuiz,getBase64Text,generateQuizFromSection } from "./generateQuiz.js";




// Rate limter 100 request per 2 hr will have to limit to user
const limiter = rateLimit({
	windowMs: 2 * HOUR, // SECOND, MINUTE, HOUR, and DAY constants are available, or a use bare number for milliseconds
	limit: 100, // Limit each IP to 100 requests per `window` (here, per 15 minutes).
	standardHeaders: 'draft-8', // draft-6: `RateLimit-*` headers; draft-7 & draft-8: combined `RateLimit` header
	legacyHeaders: false, // Disable the `X-RateLimit-*` headers.
	ipv6Subnet: 56, // Set to 60 or 64 to be less aggressive, or 52 or 48 to be more aggressive
	// store: ... , // Redis, Memcached, etc. See below.
})






const app = express();
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Private-Network", "true");
  next();
});

//Increse payload of json transfer
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ limit: '20mb', extended: true }));
app.use(limiter)


app.use(cors());
app.use(morgan("dev")); // see request in terminal
app.use(express.json());


app.post("/api/generate-quiz", async (req, res) => {
  // connect to hugging face model and geneare ques
  try {
    const { firstText, secondText, numQuestions, imageOCR } = req.body;

    if (!firstText) {
      return res.status(400).json({ error: "firstText is required" });
    }
    // console.log(firstText, secondText, numQuestions)

    const quiz = await generateQuiz(numQuestions, firstText, secondText,imageOCR);
    res.json({ quiz });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate quiz" });
  }
});


app.post("/api/generate-section-quiz", async (req, res) => {
  // connect to hugging face model and geneare ques
  try {
    const { inkAnalysisList, OCRResults, numQuestions} = req.body;



    const quiz = await generateQuizFromSection(numQuestions,inkAnalysisList,OCRResults);
    res.json({ quiz });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate quiz" });
  }
});




app.post("/api/base64-OCR", async (req,res) =>{

  try{
    const base64Strings = req.body;
    const strings = base64Strings.base64Strings
    let textPerImage =[];
    for (var i=0; i<strings.length;i++){
      let image_text = await getBase64Text(strings[i])
      console.log(image_text);
      textPerImage.push(image_text);
    }

    console.log(textPerImage[0])
    res.json({textPerImage})

  }
  catch(e){
    console.error(e)
    res.status(500).json({ error: "Failed to extract data from base64" });

  }

})


const port = process.env.PORT || 3001;
app.listen(port, () => console.log(`Server running on port ${port}`));