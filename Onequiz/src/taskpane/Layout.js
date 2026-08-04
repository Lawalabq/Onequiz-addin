

// This function updates the question and answer choices based on the current question index
export function updateQuestion(index, quizquestions) {

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














