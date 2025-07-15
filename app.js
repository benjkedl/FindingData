// Main Application Logic

/* Assessment Questions */
const assessmentQuestions = [
    {
        question: "How would you describe your experience or interest with learning new methods  with excel?",
        options: ["Not interested in new methods", "Interested, but not very experienced", "Interested, a bit experienced", "Very Interested, more experienced"]
    },
    {
        question: "How often do you use cell references in an excel sheet (like A1, B2)?",
        options: ["Never", "Rarely", "Somewhat often", "Often"]
    },
    {
        question: "Do you use Excel Functions like CONCAT() or TEXTJOIN()?",
        options: ["Never", "Rarely", "Somewhat often", "Often"]
    },
    {
        question: "Do you use Excel Shortcuts like CTRL+F?",
        options: ["Never", "Rarely", "Somewhat often", "Often"]
    },
    {
        question: "Are you interested in learning more about Excel functions to help make finding data in your work documents easier?",
        options: ["Not really", "Somewhat", "I'm interested", "I'm very interested"]
    }
];

/* Learning Pathway Data */
const basicPathway = {
    lesson1: {
        title: "Introduction to Excel Formulas",
        content: "Learn how to start calculations in Excel by using the equals sign (=). This tells Excel that you want to perform a calculation rather than just enter text. Try basic math operations like addition (+) and subtraction (-).",
        exercise: "In cell B1, enter a formula to calculate 4 + 4. Remember to start with an equals sign!",
        expectedAnswer: "=4+4",
        result: "8",
        sampleData: {}
    },
    lesson2: {
        title: "Cell References",
        content: "Instead of typing numbers directly, you can reference values in other cells. This makes your spreadsheets dynamic - if you change the original value, your formula automatically updates.",
        exercise: "In cell B2, create a formula that references the value in cell A2. Use the format =A2",
        expectedAnswer: "=A2",
        result: "John",
        sampleData: {
            "A2": "John",
            "A3": "Jane",
            "A4": "Bob"
        }
    },
    lesson3: {
        title: "Finding Data with CTRL+F",
        content: "Use CTRL+F (or Cmd+F on Mac) to quickly find specific data in your spreadsheet. This is Excel's built-in search function that highlights matching text (it also works on your browser!).",
        exercise: "Practice using CTRL+F to find the word 'POTATO' somewhere in the practice grid below. Press CTRL+F, type 'POTATO', and see it highlighted. When you've found it, click Submit to complete the basic pathway.",
        expectedAnswer: null, // No validation needed
        sampleData: {
            "A1": "Name",
            "B1": "Item",
            "C1": "Category",
            "A2": "John",
            "B2": "POTATO",
            "C2": "Vegetable",
            "A3": "Jane",
            "B3": "Carrot",
            "C3": "Vegetable"
        }
    }
};

const advancedPathway = {
    lesson1: {
        title: "Excel Functions",
        content: "Functions are pre-built formulas that perform specific calculations. They follow the format =FUNCTIONNAME(arguments). Common functions include SUM(), AVERAGE(), COUNT(), and many others.",
        exercise: "In cell B1, use the SUM function to add up the values in cells A1 through A3: =SUM(A1:A3)",
        expectedAnswer: "=SUM(A1:A3)",
        result: "15",
        sampleData: {
            "A1": "5",
            "A2": "4",
            "A3": "6"
        }
    },
    lesson2: {
        title: "XLOOKUP with Column References",
        content: "XLOOKUP is a powerful function that searches for a value in one column and returns a corresponding value from another column.  You can use this to find data in large datasets based on a piece of data you already have, like a name! Check out the format, which we will talk about on Thursday: \n Format: =XLOOKUP(lookup_value, lookup_array, return_array)",
        exercise: "Use XLOOKUP to find the email for John Doe (in cell A1). A1 is our 'Lookup_value', because we are looking up the email for John Doe.  Our Lookup_Array is the column where we expect John Doe's Name to be, which is the F column (refer to the whole column by using F:F).  And our Return_array is the column with the information we want, which is emails.  Here, they are in column G, so we will use G:G.  Our final formula: In cell B1, enter: =XLOOKUP(A1,F:F,G:G).",
        expectedAnswer: "=XLOOKUP(A1,F:F,G:G)",
        result: "JohnDoe@cuanschutz.edu",
        sampleData: {
            "A1": "John Doe",
            "F1": "Name",
            "G1": "Email",
            "F2": "John Doe",
            "G2": "JohnDoe@cuanschutz.edu",
            "F3": "Jane Doe",
            "G3": "JaneDoe@cuanschutz.edu",
            "F4": "Henry Student",
            "G4": "Henry.Student@cuanschutz.edu"
        }
    },
    lesson3: {
        title: "XLOOKUP with Table References",
        content: "Remember when we talked about turning data into Tables by using Insert>Insert table?  Well, you can use column names of tables with XLOOKUP! To refer to a column in a table, you just have to type the name of the table (EX. Table1), followed by the name of the Column you want in brackets (EX. Table1[columnName]).",
        exercise: "Use the IDTable below to find Jane's UCD Number. Just like before, we want to use the name (Jane Doe this time) as our lookup_value (A1).  We then want to find that name in the Name column of the table on the right (IDTable[Name]).  Finally, we want to find the UCD ID, which is also in the IDTable (IDTable[UCD ID]).  Our final formula: Enter into cell B1:  =XLOOKUP(A1,IDTable[Name],IDTable[UCD ID])",
        expectedAnswer: "=XLOOKUP(A1,IDTable[Name],IDTable[UCD ID])",
        result: "UCD777",
        sampleData: {
            "A1": "Jane Doe"
        },
        namedTableData: {
            "F1": "Name",
            "G1": "UCD ID",
            "F2": "John Doe",
            "G2": "UCD738492",
            "F3": "Jane Doe",
            "G3": "UCD777",
            "F4": "Henry Student",
            "G4": "UCD83910183749"
        }
    }
};

/* Page management */
const pages = {
    landing: document.getElementById('landing-page'),
    assessment: document.getElementById('assessment-page'),
    learning: document.getElementById('learning-page')
};

function showPage(page) {
    Object.values(pages).forEach(p => p.classList.remove('active'));
    pages[page].classList.add('active');
}

/* Assessment Quiz Logic */
let currentQuestionIndex = 0;
let score = 0;
const totalQuestions = assessmentQuestions.length;

const questionText = document.getElementById('question-text');
const optionsContainer = document.getElementById('options-container');
const nextBtn = document.getElementById('next-btn');
const progressFill = document.getElementById('progress-fill');
const currentQuestionSpan = document.getElementById('current-question');
const totalQuestionsSpan = document.getElementById('total-questions');
const assessmentResult = document.getElementById('assessment-result');
const resultText = document.getElementById('result-text');

function startAssessment() {
    showPage('assessment');
    currentQuestionIndex = 0;
    score = 0;
    totalQuestionsSpan.textContent = totalQuestions;
    loadQuestion();
}

function loadQuestion() {
    const current = assessmentQuestions[currentQuestionIndex];
    questionText.textContent = current.question;
    optionsContainer.innerHTML = '';
    current.options.forEach((opt, idx) => {
        const option = document.createElement('div');
        option.className = 'option';
        option.textContent = opt;
        option.onclick = () => selectOption(idx);
        optionsContainer.appendChild(option);
    });
    currentQuestionSpan.textContent = currentQuestionIndex + 1;
    nextBtn.disabled = true;
    progressFill.style.width = `${(currentQuestionIndex / totalQuestions) * 100}%`;
}

function selectOption(idx) {
    const options = optionsContainer.querySelectorAll('.option');
    options.forEach(opt => opt.classList.remove('selected'));
    options[idx].classList.add('selected');
    // Score: 0,1,2,3 respectively
    score += idx;
    nextBtn.disabled = false;
}

function nextQuestion() {
    currentQuestionIndex++;
    if (currentQuestionIndex < totalQuestions) {
        loadQuestion();
    } else {
        showAssessmentResult();
    }
}

function showAssessmentResult() {
    // Update progress width
    progressFill.style.width = '100%';

    const threshold = (totalQuestions * 3) / 2; // Half of max score
    const isAdvanced = score > threshold;
    resultText.textContent = `Based on your results, the ${isAdvanced ? 'Intermediate/Advanced' : 'Basic'} Pathway will be the most helpful for you.`;
    assessmentResult.style.display = 'block';

    // Store pathway selection
    window.selectedPathway = isAdvanced ? 'advancedPathway' : 'basicPathway';

    // Enable Start Learning button
    const startLearningBtn = document.getElementById('start-learning-btn');
    startLearningBtn.textContent = 'Start ' + (isAdvanced ? 'Intermediate/Advanced' : 'Basic') + ' Learning';
}

function startLearning() {
    showPage('learning');
    // Load the first lesson for the selected pathway
    window.currentLessonIndex = 0;
    loadLesson();
}

/* Learning Pathway Logic */
function getCurrentPathwayData() {
    return window.selectedPathway === 'advancedPathway' ? advancedPathway : basicPathway;
}

const lessonTitle = document.getElementById('lesson-title');
const lessonContent = document.getElementById('lesson-content');
const exerciseDescription = document.getElementById('exercise-description');
const excelGrid = document.getElementById('excel-grid');
const pageCounter = document.getElementById('page-counter');
const feedbackSection = document.getElementById('feedback');
const feedbackTitle = document.getElementById('feedback-title');
const feedbackMessage = document.getElementById('feedback-message');

function loadLesson() {
    const pathway = getCurrentPathwayData();
    const keys = Object.keys(pathway);
    const lessonKey = keys[window.currentLessonIndex];
    const lesson = pathway[lessonKey];

    lessonTitle.textContent = lesson.title;
    lessonContent.textContent = lesson.content;
    exerciseDescription.textContent = lesson.exercise;
    pageCounter.textContent = `${window.currentLessonIndex + 1} / ${keys.length}`;

    // Generate grid rows - use more rows for lessons with named tables
    const numRows = lesson.namedTableData ? 12 : 5;
    generateGridRows(numRows);

    // Populate sample data if any
    if (lesson.sampleData) {
        Object.entries(lesson.sampleData).forEach(([cell, value]) => {
            const cellDiv = document.getElementById(`cell-${cell}`);
            if (cellDiv) cellDiv.textContent = value;
        });
    }

    // Populate named table data if any
    if (lesson.namedTableData) {
        // Add table header styling for named table
        Object.entries(lesson.namedTableData).forEach(([cell, value]) => {
            const cellDiv = document.getElementById(`cell-${cell}`);
            if (cellDiv) {
                cellDiv.textContent = value;
                // Add table header styling to first row of named table
                if (cell.includes('1')) {
                    cellDiv.classList.add('table-header');
                }
            }
        });
    }

    // Add event listeners to all cells
    addCellEventListeners();

    // Hide feedback section
    feedbackSection.style.display = 'none';
}

function generateGridRows(numRows = 5) {
    excelGrid.innerHTML = '';
    // Column headings
    excelGrid.innerHTML += '<div></div>';
    ['A','B','C','D','E','F','G','H'].forEach(col => {
        excelGrid.innerHTML += `<div class="excel-heading">${col}</div>`;
    });

    for (let r = 1; r <= numRows; r++) {
        // Row heading
        excelGrid.innerHTML += `<div class="excel-heading">${r}</div>`;
        for (let c = 0; c < 8; c++) {
            const cellAddress = String.fromCharCode(65 + c) + r; // A1, B1, ...
            excelGrid.innerHTML += `<div class="excel-cell" id="cell-${cellAddress}" contenteditable="true"></div>`;
        }
    }
}

function addCellEventListeners() {
    const cells = document.querySelectorAll('.excel-cell');
    cells.forEach(cell => {
        // Remove existing listeners to prevent duplicates
        cell.removeEventListener('blur', handleCellBlur);
        cell.removeEventListener('keydown', handleCellKeydown);
        
        // Add new listeners
        cell.addEventListener('blur', handleCellBlur);
        cell.addEventListener('keydown', handleCellKeydown);
    });
}

function handleCellBlur(event) {
    const value = event.target.textContent.trim();
    if (value.startsWith('=')) {
        try {
            const result = evaluateFormula(value);
            event.target.textContent = result;
        } catch (err) {
            console.error('Formula error:', err);
            event.target.textContent = '#ERROR';
        }
    }
}

function handleCellKeydown(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        event.target.blur(); // This will trigger the blur handler
    }
}

function evaluateFormula(formula) {
    // Basic formula evaluation: supports +,-,*,/ and cell references like A1
    const cleaned = formula.slice(1); // Remove '='
    
    // Handle SUM function
    if (cleaned.toUpperCase().startsWith('SUM(')) {
        const match = cleaned.match(/SUM\(([A-Z]\d+):([A-Z]\d+)\)/i);
        if (match) {
            const startCell = match[1];
            const endCell = match[2];
            return calculateSum(startCell, endCell);
        }
    }
    
    // Handle XLOOKUP function (simplified simulation)
    if (cleaned.toUpperCase().includes('XLOOKUP')) {
        // For demo purposes, return the expected result
        const pathway = getCurrentPathwayData();
        const keys = Object.keys(pathway);
        const lesson = pathway[keys[window.currentLessonIndex]];
        return lesson.result || '$1.50';
    }
    
    // Replace cell references with their values
const replaced = cleaned.replace(/[A-Z]\d+/g, match => {
    const cellDiv = document.getElementById(`cell-${match}`);
    if (!cellDiv) return '0'; // If cell not found, treat as 0
    const value = cellDiv.textContent.trim();
    if (value === '') return '0'; // Empty cell treated as 0
    // If value is numeric, return it; otherwise, wrap it in quotes for safe evaluation
    return isNaN(value) ? `"${value}"` : value;
});

    
    // Use Function constructor for math evaluation (safe for basic expressions)
    try {
        return new Function(`return ${replaced}`)();
    } catch (e) {
        return '#ERROR';
    }
}

function calculateSum(startCell, endCell) {
    // Extract column and row from cell references
    const startCol = startCell.match(/[A-Z]/)[0];
    const startRow = parseInt(startCell.match(/\d+/)[0]);
    const endCol = endCell.match(/[A-Z]/)[0];
    const endRow = parseInt(endCell.match(/\d+/)[0]);
    
    let sum = 0;
    for (let row = startRow; row <= endRow; row++) {
        const cellId = `cell-${startCol}${row}`;
        const cellDiv = document.getElementById(cellId);
        if (cellDiv) {
            const value = parseFloat(cellDiv.textContent.trim()) || 0;
            sum += value;
        }
    }
    return sum;
}

/* Submit Logic */
function submitAnswer() {
    const pathway = getCurrentPathwayData();
    const keys = Object.keys(pathway);
    const lesson = pathway[keys[window.currentLessonIndex]];

    // For CTRL+F lesson (lesson3 of basic pathway), just proceed
    if (window.selectedPathway === 'basicPathway' && window.currentLessonIndex === 2) {
        showFeedback(true, 'CTRL+F Exercise Complete', 'Great job! You have completed the basic pathway.');
        return;
    }

    // Check expected answer for other lessons
    if (lesson.expectedAnswer) {
        // Look for any cell that contains the expected formula
        const cells = document.querySelectorAll('.excel-cell');
        let foundCorrectAnswer = false;
        
        cells.forEach(cell => {
            const value = cell.textContent.trim();
            if (value === lesson.expectedAnswer || value === lesson.result) {
                foundCorrectAnswer = true;
            }
        });

        if (foundCorrectAnswer) {
            showFeedback(true, 'Correct!', 'Well done! You can proceed to the next lesson.');
        } else {
            showFeedback(false, 'Try Again', 'Please check your formula and try again.');
        }
    }
}

function showFeedback(isSuccess, title, message) {
    feedbackSection.className = 'feedback-section ' + (isSuccess ? 'success' : 'error');
    feedbackTitle.textContent = title;
    feedbackMessage.textContent = message;
    feedbackSection.style.display = 'block';
}

function resetExercise() {
    // Reload lesson
    loadLesson();
}

/* Navigation */
function nextPage() {
    const pathway = getCurrentPathwayData();
    const keys = Object.keys(pathway);
    if (window.currentLessonIndex < keys.length - 1) {
        window.currentLessonIndex++;
        loadLesson();
    } else {
        // Completed all lessons
        showCompletionMessage();
    }
}

function previousPage() {
    if (window.currentLessonIndex > 0) {
        window.currentLessonIndex--;
        loadLesson();
    }
}

function showCompletionMessage() {
    const pathwayName = window.selectedPathway === 'advancedPathway' ? 'Intermediate/Advanced' : 'Basic';
    lessonTitle.textContent = 'Congratulations!';
    lessonContent.textContent = `You have successfully completed the ${pathwayName} pathway. You now have stronger Excel skills for finding and manipulating data.`;
    exerciseDescription.textContent = 'Click "Main Menu" to return to the start or review previous lessons.';
    excelGrid.innerHTML = '<div class="completion-message">🎉 Learning Complete! 🎉</div>';
    feedbackSection.style.display = 'none';
}

function goToMenu() {
    showPage('landing');
}

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    // Show landing page by default
    showPage('landing');
});