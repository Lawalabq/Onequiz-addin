/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */




Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();

      const firstTextPrediction = await getInkAnalysisResults(context, page, 0);
      const secondTextPrediction = await getInkAnalysisResults(context, page, 1);
      
      const res = await fetch("api/generate-quiz", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ firstText: firstTextPrediction, secondText: secondTextPrediction, numQuestions: 10 })
    });

    if (res.ok){
      let { quiz } = await res.json();
      
      quiz=JSON.parse(quiz);

    

      // Display the quiz questions in the task pane
      const quizquestions = quiz.questions;
      const quizContainer = document.getElementById("quiz-form");
      quizContainer.style.display = "block";

      // This populates the question and answer choices in the task pane
      let currentQuestionIndex = 0;
      updateQuestion(currentQuestionIndex, quizquestions);


    // Adding event listeners to the navigation buttons
    let prevButton = document.getElementById("prev-question");
    let nextButton = document.getElementById("next-question");

    prevButton.addEventListener("click", function(event) {
      event.preventDefault();
      if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        updateQuestion(currentQuestionIndex, quizquestions);
      }
    });

    nextButton.addEventListener("click", function(event) {
      event.preventDefault();
      if (currentQuestionIndex < quizquestions.length - 1) {
        currentQuestionIndex++;
        updateQuestion(currentQuestionIndex, quizquestions);
      }
    });



  }});





    

      
  
  } catch (error) {
    console.error(error);
  }
}













// This function updates the question and answer choices based on the current question index
function updateQuestion(index, quizquestions) {

  // remove event lsistener from the check answer button to avoid multiple event listeners being added
  let checkAnswerButton = document.getElementById("check-answer");
  let newCheckAnswerButton = checkAnswerButton.cloneNode(true);
  checkAnswerButton.parentNode.replaceChild(newCheckAnswerButton, checkAnswerButton);


  //Setting the current question index and displaying the first question
      let currentQuestionIndex = index;
      let questionNumberElement = document.getElementById("question-number");
      questionNumberElement.textContent = `Question ${currentQuestionIndex + 1} of ${quizquestions.length}`;

      //clear the previous answer choices
      let answerElement = document.getElementById("question_answer");
      answerElement.innerHTML = "";


      let questionElement = document.getElementById("question");
      questionElement.textContent = quizquestions[currentQuestionIndex].question;

      let questionType = quizquestions[currentQuestionIndex].questionType;

      // Displaying the answer choices based on the question type

      //this needs effecienyc work like the double storing of options and the answer
      if (questionType === "multiple_choice") {
        let answerElement = document.getElementById("question_answer");
        for ( let choice in quizquestions[currentQuestionIndex].choices) {
          let choiceElement = document.createElement("input");
          choiceElement.type = "radio";
          choiceElement.name = "answer";
          choiceElement.value = quizquestions[currentQuestionIndex].choices[choice];
          answerElement.appendChild(choiceElement);

          let labelElement = document.createElement("label");
          labelElement.textContent = quizquestions[currentQuestionIndex].choices[choice];
          answerElement.appendChild(labelElement);

          
        }

        // Adding event listener to the check answer button
        let checkAnswerButton = document.getElementById("check-answer");
        checkAnswerButton.addEventListener("click", function(event) {
          event.preventDefault();
         
        let selectedAnswer = document.querySelector('input[name="answer"]:checked');
         console.log("selected ",selectedAnswer.value);
         console.log("answer ", quizquestions[currentQuestionIndex].correctAnswer);

        if (selectedAnswer.value === quizquestions[currentQuestionIndex].correctAnswer) {
          let labelElement = selectedAnswer.nextSibling;
          labelElement.style.color = "green";
        }
        else {
          let labelElement = selectedAnswer.nextSibling;
          labelElement.style.color = "red";
        }})




      }
      else if (questionType === "short_answer") {
        let answerElement = document.getElementById("question_answer");
        let textInputElement = document.createElement("input");
        textInputElement.type = "text";
        textInputElement.name = "answer";
        answerElement.appendChild(textInputElement);

        // Adding event listener to the check answer button
        let checkAnswerButton = document.getElementById("check-answer");
        checkAnswerButton.addEventListener("click", function(event) {
        event.preventDefault();
        let userAnswer = textInputElement.value.trim().toLowerCase();
        
        let correctAnswer = quizquestions[currentQuestionIndex].correctAnswer.trim().toLowerCase();

        if (correctAnswer.includes(userAnswer)) {
          textInputElement.style.backgroundColor = "green";
        }
        else {
          textInputElement.style.backgroundColor = "red";
        }})
      }

    }








async function getInkAnalysisResults(context, page, wordAlternatesChoice = 0) {
  const inkAnalysis = page.inkAnalysisOrNull;
  inkAnalysis.load("id");
  await context.sync();

  if (inkAnalysis.isNull) {
    console.log("No ink analysis available.");
    return "";
  }

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