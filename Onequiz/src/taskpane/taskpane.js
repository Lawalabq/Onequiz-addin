/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */


import {getInkAnalysisResults,getImageOcrData} from "./oneNoteExtraction.js"
import { updateQuestion } from "./Layout.js";


// Sign In Logic
// async function getUserToken() {
//   try {
//     const accessToken = await OfficeRuntime.auth.getAccessToken({
//       allowSignInPrompt: true,
//       allowConsentPrompt: true,
//       forMSGraphAccess: false
//     });
//     return accessToken;
//   } catch (error) {
//     console.error("SSO token error:", error.code, error.message);
//     throw error;
//   }
// }


Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("test_run").onclick = test_run;
  }

// Sign In Logic

  if (info.host === Office.HostType.OneNote) {
    try {
      userAccessToken = await getUserToken();
      console.log("Signed in successfully");
      // now you can call your backend, e.g. loadUserData();
    } catch (error) {
      console.error("Failed to get SSO token on load:", error);
      // fall back to a manual "Sign in" button here (see below)
    }
  }



});

export async function run() {
  /* Main functin for generating quizes from pages

  */
  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
    

      const firstTextPrediction = await getInkAnalysisResults(context, page, 0);
      const secondTextPrediction = await getInkAnalysisResults(context, page, 1);
      const imageOCR = await getImageOcrData(context,page);
      
      const res = await fetch("api/generate-quiz", {
      method: "POST",
      headers: { "Content-Type": "application/json",'Authorization': `Bearer ${userAccessToken}` },
      body: JSON.stringify({ firstText: firstTextPrediction, secondText: secondTextPrediction, imageText:imageOCR, numQuestions: 10 })
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



  }
else if (response.status === 401){
      userAccessToken = await getUserToken();
    }

});

  
  } catch (error) {
    console.error(error);
  }
}







export async function test_run() {
  // Failed function pages need to be active to get info from them
  try {
    await OneNote.run(async (context) => {  
       const pages = context.application.getActiveSection().pages;
       pages.load('items');
       await context.sync();

       let currentActivePage = context.application.getActivePage();

       let pageInkAnalysisList = [];
       let pageOCRText =[];
      // console.log(pages);
       for (var i =0 ; i<pages.items.length;i++){
        const page =pages.items[i];
        // console.log(page.inkAnalysisOrNull);
        context.application.navigateToPage(page);
        let inkAnalysisResults = await getInkAnalysisResults(context,page,0);
        let pageOCRResults= await getImageOcrData(context,page);
        pageInkAnalysisList.push(inkAnalysisResults);
        pageOCRText.push(pageOCRResults);

       }
       await context.sync();

       //Need to add question number selection

      context.application.navigateToPage(currentActivePage);
      const res = await fetch("api/generate-section-quiz", {
      method: "POST",
      headers: { "Content-Type": "application/json" ,'Authorization': `Bearer ${userAccessToken}`},
      body: JSON.stringify({ inkAnalysisList:pageInkAnalysisList, OCRResults:pageOCRText ,numQuestions: 20 })
    });

    if (res.ok){

      let {quiz} = await res.json();
      quiz = JSON.parse(quiz);
      startQuiz(quiz);
    }
    else if (response.status === 401){
      userAccessToken = await getUserToken();
    }



      

       // 
       
       

    










    //   const res = await fetch("api/base64-OCR", {
    //   method: "POST",
    //   headers: { "Content-Type": "application/json" },
    //   body: JSON.stringify({ base64Strings: base64Strings})
    // });

    


    
       


    })
  }
  catch (error) {
    console.error(error);
  }};





// export async function test_run() {
//   try {
//     await OneNote.run(async (context) => {  
//        const page = context.application.getActivePage();
//        let image_json= await getImageOcrData(context,page);

//        const base64Strings = image_json.map(u => u.base64);
//        const onenoteOCR = image_json.map(u=> u.ocrText);
//        console.log(onenoteOCR);



//     //   const res = await fetch("api/base64-OCR", {
//     //   method: "POST",
//     //   headers: { "Content-Type": "application/json" },
//     //   body: JSON.stringify({ base64Strings: base64Strings})
//     // });

//     if (res.ok){
//       let imageText = await res.json;
//       console.log(imageText)
//     }


    
       


//     })
//   }
//   catch (error) {
//     console.error(error);
//   }};


function startQuiz(quiz){
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



  };

































