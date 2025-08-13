// Application Data from JSON
const appData = {
  "topicCategories": [
    {
      "level": "Beginner",
      "color": "#4ade80",
      "topics": [
        {
          "id": "B1",
          "title": "Excel Basics & Navigation",
          "description": "Learn Excel interface, workbooks, worksheets, and basic navigation",
          "duration": "30 min",
          "exercises": 5,
          "keyPoints": ["Interface overview", "Workbook vs Worksheet", "Cell navigation", "Ribbon and tabs"],
          "sampleData": "Basic employee list",
          "difficulty": 1
        },
        {
          "id": "B2", 
          "title": "Data Entry & Formatting",
          "description": "Master data input, cell formatting, and number formats",
          "duration": "45 min",
          "exercises": 7,
          "keyPoints": ["Data types", "Cell formatting", "Number formats", "Font styles"],
          "sampleData": "Product catalog",
          "difficulty": 1
        },
        {
          "id": "B3",
          "title": "Basic Formulas",
          "description": "Introduction to formulas and basic calculations",
          "duration": "40 min", 
          "exercises": 6,
          "keyPoints": ["Formula syntax", "SUM", "AVERAGE", "COUNT", "Basic operators"],
          "sampleData": "Sales figures",
          "difficulty": 2
        },
        {
          "id": "B4",
          "title": "Cell References",
          "description": "Understanding relative, absolute, and mixed references",
          "duration": "35 min",
          "exercises": 5,
          "keyPoints": ["A1 vs $A$1", "Relative references", "Absolute references", "Mixed references"],
          "sampleData": "Budget template",
          "difficulty": 2
        },
        {
          "id": "B5",
          "title": "Basic Functions",
          "description": "Essential Excel functions for everyday use",
          "duration": "50 min",
          "exercises": 8,
          "keyPoints": ["MAX", "MIN", "TODAY", "TEXT functions", "Function syntax"],
          "sampleData": "Student grades",
          "difficulty": 2
        },
        {
          "id": "B6",
          "title": "Simple Charts",
          "description": "Create and format basic charts and graphs",
          "duration": "40 min",
          "exercises": 6,
          "keyPoints": ["Chart types", "Data selection", "Chart formatting", "Bar and pie charts"],
          "sampleData": "Survey results",
          "difficulty": 2
        }
      ]
    },
    {
      "level": "Intermediate",
      "color": "#f59e0b",
      "topics": [
        {
          "id": "I1",
          "title": "Advanced Formulas",
          "description": "Master IF statements and logical functions",
          "duration": "60 min",
          "exercises": 10,
          "keyPoints": ["IF function", "AND/OR logic", "Nested IF", "IFERROR"],
          "sampleData": "Commission calculator",
          "difficulty": 3
        },
        {
          "id": "I2",
          "title": "Lookup Functions", 
          "description": "VLOOKUP, HLOOKUP, and INDEX-MATCH mastery",
          "duration": "75 min",
          "exercises": 12,
          "keyPoints": ["VLOOKUP syntax", "HLOOKUP", "INDEX-MATCH", "Approximate vs exact match"],
          "sampleData": "Customer database",
          "difficulty": 4
        },
        {
          "id": "I3",
          "title": "Text Functions",
          "description": "Manipulate and clean text data effectively",
          "duration": "55 min",
          "exercises": 9,
          "keyPoints": ["LEFT/RIGHT/MID", "CONCATENATE", "UPPER/LOWER", "SUBSTITUTE"],
          "sampleData": "Contact list cleanup",
          "difficulty": 3
        },
        {
          "id": "I4",
          "title": "Date & Time Functions",
          "description": "Work with dates, times, and duration calculations",
          "duration": "65 min",
          "exercises": 11,
          "keyPoints": ["DATE function", "DATEDIF", "WEEKDAY", "Time calculations"],
          "sampleData": "Project timeline",
          "difficulty": 3
        },
        {
          "id": "I5",
          "title": "Conditional Formatting",
          "description": "Highlight data based on conditions and rules",
          "duration": "45 min",
          "exercises": 7,
          "keyPoints": ["Color scales", "Data bars", "Icon sets", "Custom rules"],
          "sampleData": "Performance metrics",
          "difficulty": 3
        },
        {
          "id": "I6",
          "title": "Data Validation",
          "description": "Control data input with validation rules",
          "duration": "40 min",
          "exercises": 6,
          "keyPoints": ["Dropdown lists", "Input restrictions", "Error messages", "Custom validation"],
          "sampleData": "Order form",
          "difficulty": 3
        },
        {
          "id": "I7",
          "title": "Advanced Charts",
          "description": "Create sophisticated visualizations",
          "duration": "70 min",
          "exercises": 10,
          "keyPoints": ["Line charts", "Scatter plots", "Combo charts", "Chart customization"],
          "sampleData": "Financial trends",
          "difficulty": 4
        },
        {
          "id": "I8",
          "title": "Basic Pivot Tables",
          "description": "Summarize and analyze data with pivot tables",
          "duration": "80 min",
          "exercises": 12,
          "keyPoints": ["Pivot table creation", "Field arrangement", "Value calculations", "Formatting"],
          "sampleData": "Sales analysis",
          "difficulty": 4
        }
      ]
    },
    {
      "level": "Advanced", 
      "color": "#ef4444",
      "topics": [
        {
          "id": "A1",
          "title": "Advanced Pivot Tables",
          "description": "Master complex pivot table features",
          "duration": "90 min",
          "exercises": 15,
          "keyPoints": ["Calculated fields", "Slicers", "Pivot charts", "Multiple consolidation"],
          "sampleData": "Multi-dimensional sales data",
          "difficulty": 5
        },
        {
          "id": "A2",
          "title": "Data Analysis Tools",
          "description": "Goal Seek, Solver, and What-If Analysis",
          "duration": "85 min",
          "exercises": 12,
          "keyPoints": ["Goal Seek", "Solver", "Data tables", "Scenario manager"],
          "sampleData": "Financial modeling",
          "difficulty": 5
        },
        {
          "id": "A3",
          "title": "Array Formulas",
          "description": "Complex calculations with array formulas",
          "duration": "75 min",
          "exercises": 10,
          "keyPoints": ["Array syntax", "Dynamic arrays", "SUMPRODUCT", "Complex criteria"],
          "sampleData": "Advanced analytics",
          "difficulty": 5
        },
        {
          "id": "A4",
          "title": "Advanced Lookup Functions",
          "description": "XLOOKUP and complex lookup scenarios",
          "duration": "70 min",
          "exercises": 11,
          "keyPoints": ["XLOOKUP", "Multiple criteria", "Two-way lookups", "Error handling"],
          "sampleData": "Complex database",
          "difficulty": 5
        },
        {
          "id": "A5",
          "title": "Dashboard Creation",
          "description": "Build interactive business dashboards",
          "duration": "100 min",
          "exercises": 8,
          "keyPoints": ["Dashboard design", "Interactive elements", "Data connections", "KPI visualization"],
          "sampleData": "Executive dashboard data",
          "difficulty": 5
        },
        {
          "id": "A6",
          "title": "Automation Basics",
          "description": "Introduction to macros and VBA",
          "duration": "95 min",
          "exercises": 10,
          "keyPoints": ["Macro recording", "Basic VBA", "Automation workflows", "Custom functions"],
          "sampleData": "Repetitive task examples",
          "difficulty": 5
        }
      ]
    }
  ],
  "sampleDatasets": {
    "sales": [
      ["Month", "Region", "Product", "Sales", "Target"],
      ["Jan", "North", "Laptop", 15000, 12000],
      ["Jan", "South", "Laptop", 12000, 10000],
      ["Jan", "East", "Desktop", 18000, 15000],
      ["Feb", "North", "Laptop", 16500, 12000],
      ["Feb", "South", "Desktop", 14000, 13000]
    ],
    "employees": [
      ["ID", "Name", "Department", "Salary", "Hire Date", "Performance"],
      [1001, "John Smith", "Sales", 50000, "2020-01-15", "Excellent"],
      [1002, "Jane Doe", "Marketing", 55000, "2019-03-20", "Good"],
      [1003, "Bob Johnson", "IT", 65000, "2018-07-10", "Excellent"],
      [1004, "Alice Brown", "HR", 52000, "2021-02-01", "Good"]
    ],
    "products": [
      ["Product ID", "Product Name", "Category", "Price", "Stock", "Supplier"],
      ["P001", "Wireless Mouse", "Electronics", 25.99, 150, "TechCorp"],
      ["P002", "USB Cable", "Electronics", 12.50, 300, "TechCorp"],
      ["P003", "Notebook", "Stationery", 3.25, 500, "OfficeMax"],
      ["P004", "Pen Set", "Stationery", 8.99, 200, "OfficeMax"]
    ]
  },
  "achievements": [
    {"id": "beginner", "name": "Excel Explorer", "description": "Complete first beginner topic", "icon": "🌟"},
    {"id": "formulas", "name": "Formula Master", "description": "Complete all basic formula exercises", "icon": "🔢"},
    {"id": "charts", "name": "Chart Champion", "description": "Create 5 different chart types", "icon": "📊"},
    {"id": "pivot", "name": "Pivot Pro", "description": "Master pivot table basics", "icon": "🔄"},
    {"id": "advanced", "name": "Excel Expert", "description": "Complete an advanced topic", "icon": "🏆"},
    {"id": "speed", "name": "Speed Demon", "description": "Complete exercise in under 2 minutes", "icon": "⚡"},
    {"id": "perfect", "name": "Perfect Score", "description": "Get 100% on any quiz", "icon": "💎"},
    {"id": "streak", "name": "Learning Streak", "description": "Study for 7 consecutive days", "icon": "🔥"}
  ]
};

// Excel Functions Library
const excelFunctions = {
  // Math Functions
  'SUM': { description: 'Adds numbers in a range', syntax: '=SUM(range)', category: 'Math' },
  'AVERAGE': { description: 'Calculates the average of numbers', syntax: '=AVERAGE(range)', category: 'Math' },
  'COUNT': { description: 'Counts cells containing numbers', syntax: '=COUNT(range)', category: 'Math' },
  'MAX': { description: 'Returns the largest value', syntax: '=MAX(range)', category: 'Math' },
  'MIN': { description: 'Returns the smallest value', syntax: '=MIN(range)', category: 'Math' },
  'ROUND': { description: 'Rounds a number to specified digits', syntax: '=ROUND(number, digits)', category: 'Math' },
  
  // Logical Functions
  'IF': { description: 'Returns one value if condition is true, another if false', syntax: '=IF(condition, true_value, false_value)', category: 'Logical' },
  'AND': { description: 'Returns TRUE if all conditions are true', syntax: '=AND(condition1, condition2, ...)', category: 'Logical' },
  'OR': { description: 'Returns TRUE if any condition is true', syntax: '=OR(condition1, condition2, ...)', category: 'Logical' },
  
  // Lookup Functions
  'VLOOKUP': { description: 'Looks up a value in the first column and returns a value in the same row', syntax: '=VLOOKUP(lookup_value, table, column, exact)', category: 'Lookup' },
  'INDEX': { description: 'Returns a value from a specific position in a range', syntax: '=INDEX(array, row, column)', category: 'Lookup' },
  'MATCH': { description: 'Returns the position of a value in a range', syntax: '=MATCH(lookup_value, array, match_type)', category: 'Lookup' },
  
  // Text Functions
  'LEFT': { description: 'Returns characters from the left side of text', syntax: '=LEFT(text, num_chars)', category: 'Text' },
  'RIGHT': { description: 'Returns characters from the right side of text', syntax: '=RIGHT(text, num_chars)', category: 'Text' },
  'MID': { description: 'Returns characters from the middle of text', syntax: '=MID(text, start, length)', category: 'Text' },
  'CONCATENATE': { description: 'Joins text strings together', syntax: '=CONCATENATE(text1, text2, ...)', category: 'Text' },
  'UPPER': { description: 'Converts text to uppercase', syntax: '=UPPER(text)', category: 'Text' },
  'LOWER': { description: 'Converts text to lowercase', syntax: '=LOWER(text)', category: 'Text' },
  
  // Date Functions
  'TODAY': { description: 'Returns the current date', syntax: '=TODAY()', category: 'Date' },
  'NOW': { description: 'Returns the current date and time', syntax: '=NOW()', category: 'Date' },
  'DATE': { description: 'Creates a date from year, month, and day', syntax: '=DATE(year, month, day)', category: 'Date' },
  'YEAR': { description: 'Returns the year from a date', syntax: '=YEAR(date)', category: 'Date' },
  'MONTH': { description: 'Returns the month from a date', syntax: '=MONTH(date)', category: 'Date' },
  'DAY': { description: 'Returns the day from a date', syntax: '=DAY(date)', category: 'Date' }
};

// Application State Management
class ExcelLearningApp {
  constructor() {
    this.currentView = 'welcome';
    this.currentTopic = null;
    this.currentLesson = 0;
    this.currentQuiz = null;
    this.currentQuestionIndex = 0;
    this.selectedLevel = 'all';
    this.spreadsheetData = [];
    this.practiceSpreadsheetData = [];
    this.selectedCell = { row: 0, col: 0 };
    this.selectedPracticeCell = { row: 0, col: 0 };
    this.undoStack = [];
    this.redoStack = [];
    this.practiceUndoStack = [];
    this.practiceRedoStack = [];
    this.chart = null;
    this.userProgress = {
      completedTopics: [],
      completedLessons: [],
      earnedAchievements: [],
      quizScores: {},
      totalProgress: 0
    };
    
    console.log('ExcelLearningApp constructor called');
    this.init();
  }

  init() {
    console.log('Initializing ExcelMaster Pro...');
    this.initializeSpreadsheets();
    this.renderTopics();
    this.renderAchievements();
    this.renderFunctions();
    this.updateProgress();
    this.bindEvents();
    console.log('App initialized successfully');
  }

  // Initialize spreadsheets
  initializeSpreadsheets() {
    // Main spreadsheet
    this.spreadsheetData = Array(30).fill().map(() => Array(15).fill(''));
    this.renderSpreadsheet();
    
    // Practice spreadsheet
    this.practiceSpreadsheetData = Array(30).fill().map(() => Array(15).fill(''));
    this.renderPracticeSpreadsheet();
  }

  // Event binding - Fixed version
  bindEvents() {
    console.log('Binding events...');

    // Use event delegation for all clicks
    document.addEventListener('click', (e) => {
      console.log('Click detected on:', e.target);
      this.handleClick(e);
    });

    // Input events
    document.addEventListener('input', (e) => {
      this.handleInput(e);
    });

    // Keyboard events
    document.addEventListener('keydown', (e) => {
      this.handleKeydown(e);
    });

    // Prevent form submissions
    document.addEventListener('submit', (e) => {
      e.preventDefault();
    });

    console.log('Events bound successfully');
  }

  handleClick(e) {
    const target = e.target;
    const closest = (selector) => target.closest(selector);

    console.log('Handling click on:', target.tagName, target.className, target.id);

    // Level buttons
    if (closest('.level-btn')) {
      e.preventDefault();
      const level = target.getAttribute('data-level');
      console.log('Level button clicked:', level);
      this.filterTopics(level);
      return;
    }

    // Topic cards
    if (closest('.topic-card')) {
      e.preventDefault();
      const topicId = closest('.topic-card').getAttribute('data-topic-id');
      console.log('Topic card clicked:', topicId);
      this.openTopic(topicId);
      return;
    }

    // Back to topics
    if (target.id === 'backToTopicsBtn') {
      e.preventDefault();
      console.log('Back button clicked');
      this.showWelcomeScreen();
      return;
    }

    // Tab buttons
    if (closest('.tab-btn')) {
      e.preventDefault();
      const tabName = target.getAttribute('data-tab');
      console.log('Tab clicked:', tabName);
      this.switchTab(tabName);
      return;
    }

    // Lesson navigation
    if (target.id === 'nextLessonBtn') {
      e.preventDefault();
      console.log('Next lesson clicked');
      this.nextLesson();
      return;
    }
    if (target.id === 'prevLessonBtn') {
      e.preventDefault();
      console.log('Previous lesson clicked');
      this.prevLesson();
      return;
    }

    // Quiz navigation
    if (target.id === 'nextQuestionBtn') {
      e.preventDefault();
      this.nextQuestion();
      return;
    }
    if (target.id === 'prevQuestionBtn') {
      e.preventDefault();
      this.prevQuestion();
      return;
    }
    if (target.id === 'submitQuizBtn') {
      e.preventDefault();
      this.submitQuiz();
      return;
    }
    if (target.id === 'completeTopicBtn') {
      e.preventDefault();
      this.completeTopic();
      return;
    }

    // Answer options
    if (closest('.answer-option')) {
      e.preventDefault();
      this.selectAnswer(target);
      return;
    }

    // Spreadsheet cells
    if (closest('td[data-row]')) {
      const row = parseInt(closest('td[data-row]').getAttribute('data-row'));
      const col = parseInt(closest('td[data-row]').getAttribute('data-col'));
      this.selectCell(row, col);
      return;
    }

    // Practice spreadsheet cells
    if (closest('td[data-practice-row]')) {
      const row = parseInt(closest('td[data-practice-row]').getAttribute('data-practice-row'));
      const col = parseInt(closest('td[data-practice-row]').getAttribute('data-practice-col'));
      this.selectPracticeCell(row, col);
      return;
    }

    // Modal triggers
    if (target.id === 'aiHelperBtn') {
      e.preventDefault();
      console.log('AI Helper button clicked');
      this.showModal('aiHelper');
      return;
    }
    if (target.id === 'achievementsBtn') {
      e.preventDefault();
      console.log('Achievements button clicked');
      this.showModal('achievements');
      return;
    }
    if (target.id === 'chartsBtn') {
      e.preventDefault();
      console.log('Charts button clicked');
      this.showModal('charts');
      return;
    }
    if (target.id === 'functionsBtn') {
      e.preventDefault();
      console.log('Functions button clicked');
      this.showModal('functions');
      return;
    }
    if (target.id === 'practiceBtn') {
      e.preventDefault();
      console.log('Practice button clicked');
      this.showModal('practiceMode');
      return;
    }

    // Modal controls
    if (closest('.modal-overlay') || closest('.modal-close')) {
      e.preventDefault();
      const modal = closest('.modal');
      if (modal) {
        const modalName = modal.id.replace('Modal', '');
        console.log('Closing modal:', modalName);
        this.hideModal(modalName);
      }
      return;
    }

    // Data loading
    if (closest('.data-btn')) {
      e.preventDefault();
      const dataset = target.getAttribute('data-dataset');
      console.log('Data button clicked:', dataset);
      this.loadSampleData(dataset);
      return;
    }

    // Spreadsheet controls
    if (target.id === 'executeFormulaBtn') {
      e.preventDefault();
      this.executeFormula();
      return;
    }
    if (target.id === 'executePracticeFormulaBtn') {
      e.preventDefault();
      this.executePracticeFormula();
      return;
    }
    if (target.id === 'undoBtn') {
      e.preventDefault();
      this.undo();
      return;
    }
    if (target.id === 'redoBtn') {
      e.preventDefault();
      this.redo();
      return;
    }
    if (target.id === 'clearBtn') {
      e.preventDefault();
      this.clearSpreadsheet();
      return;
    }
    if (target.id === 'saveBtn') {
      e.preventDefault();
      this.save();
      return;
    }

    // Practice mode controls
    if (target.id === 'loadPracticeDataBtn') {
      e.preventDefault();
      this.loadPracticeData();
      return;
    }

    // Chat
    if (target.id === 'sendChatBtn') {
      e.preventDefault();
      this.sendChatMessage();
      return;
    }

    // Charts
    if (target.id === 'createChartBtn') {
      e.preventDefault();
      this.createChart();
      return;
    }

    // Notifications
    if (target.id === 'notificationClose') {
      e.preventDefault();
      this.hideNotification();
      return;
    }

    // Function items
    if (closest('.function-item')) {
      e.preventDefault();
      const funcName = closest('.function-item').getAttribute('data-function');
      this.insertFunction(funcName);
      return;
    }

    // Exercise items
    if (closest('.exercise-item')) {
      e.preventDefault();
      const exercise = closest('.exercise-item').getAttribute('data-exercise');
      this.startExercise(exercise);
      return;
    }
  }

  handleInput(e) {
    // Functions search
    if (e.target.id === 'functionsSearch') {
      this.filterFunctions(e.target.value);
      return;
    }
  }

  handleKeydown(e) {
    // Chat input
    if (e.target.id === 'chatInput' && e.key === 'Enter') {
      e.preventDefault();
      this.sendChatMessage();
      return;
    }

    // Formula inputs
    if ((e.target.id === 'formulaInput' || e.target.id === 'practiceFormulaInput') && e.key === 'Enter') {
      e.preventDefault();
      if (e.target.id === 'formulaInput') {
        this.executeFormula();
      } else {
        this.executePracticeFormula();
      }
      return;
    }

    // Keyboard shortcuts
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'z':
          e.preventDefault();
          this.undo();
          break;
        case 'y':
          e.preventDefault();
          this.redo();
          break;
        case 's':
          e.preventDefault();
          this.save();
          break;
      }
    }

    // Escape key to close modals
    if (e.key === 'Escape') {
      const openModal = document.querySelector('.modal:not(.hidden)');
      if (openModal) {
        this.hideModal(openModal.id.replace('Modal', ''));
      }
    }
  }

  // Topic Management - Fixed version
  renderTopics() {
    console.log('Rendering topics...');
    const topicsGrid = document.getElementById('topicsGrid');
    if (!topicsGrid) {
      console.error('Topics grid not found');
      return;
    }

    const allTopics = appData.topicCategories.flatMap(category => 
      category.topics.map(topic => ({...topic, level: category.level, color: category.color}))
    );

    const filteredTopics = this.selectedLevel === 'all' 
      ? allTopics 
      : allTopics.filter(topic => topic.level === this.selectedLevel);

    console.log('Filtered topics:', filteredTopics.length);

    topicsGrid.innerHTML = filteredTopics.map(topic => `
      <div class="topic-card ${this.userProgress.completedTopics.includes(topic.id) ? 'completed' : ''}" 
           data-topic-id="${topic.id}">
        <div class="topic-header">
          <span class="topic-id">${topic.id}</span>
          <span class="topic-status">
            ${this.userProgress.completedTopics.includes(topic.id) ? '✅' : '📖'}
          </span>
        </div>
        <h3>${topic.title}</h3>
        <p class="topic-description">${topic.description}</p>
        <div class="topic-meta">
          <div class="topic-tags">
            <span class="difficulty-badge" data-difficulty="${topic.difficulty}">
              ${topic.difficulty <= 2 ? 'Beginner' : topic.difficulty <= 4 ? 'Intermediate' : 'Advanced'}
            </span>
          </div>
          <div class="topic-details">
            <span>${topic.duration}</span> • <span>${topic.exercises} exercises</span>
          </div>
        </div>
      </div>
    `).join('');

    // Update statistics
    const totalTopics = allTopics.length;
    const completedTopics = this.userProgress.completedTopics.length;
    const earnedBadges = this.userProgress.earnedAchievements.length;

    const totalTopicsEl = document.getElementById('totalTopics');
    const completedTopicsEl = document.getElementById('completedTopics');
    const earnedBadgesEl = document.getElementById('earnedBadges');

    if (totalTopicsEl) totalTopicsEl.textContent = totalTopics;
    if (completedTopicsEl) completedTopicsEl.textContent = completedTopics;
    if (earnedBadgesEl) earnedBadgesEl.textContent = earnedBadges;

    console.log('Topics rendered successfully');
  }

  filterTopics(level) {
    console.log('Filtering topics by level:', level);
    this.selectedLevel = level;
    
    // Update active level button
    document.querySelectorAll('.level-btn').forEach(btn => btn.classList.remove('active'));
    const activeBtn = document.querySelector(`[data-level="${level}"]`);
    if (activeBtn) {
      activeBtn.classList.add('active');
    }
    
    this.renderTopics();
    this.showNotification(`Filtered topics: ${level === 'all' ? 'All Levels' : level}`);
  }

  openTopic(topicId) {
    console.log('Opening topic:', topicId);
    const allTopics = appData.topicCategories.flatMap(category => 
      category.topics.map(topic => ({...topic, level: category.level}))
    );
    
    this.currentTopic = allTopics.find(topic => topic.id === topicId);
    if (!this.currentTopic) {
      console.error('Topic not found:', topicId);
      return;
    }

    console.log('Found topic:', this.currentTopic.title);
    this.currentLesson = 0;
    this.showTopicDetailView();
    this.renderTopicContent();
    this.showNotification(`Started topic: ${this.currentTopic.title}`);
  }

  showTopicDetailView() {
    console.log('Showing topic detail view');
    const welcomeScreen = document.getElementById('welcomeScreen');
    const topicDetailView = document.getElementById('topicDetailView');

    if (welcomeScreen) {
      welcomeScreen.classList.add('hidden');
    }
    if (topicDetailView) {
      topicDetailView.classList.remove('hidden');
    }

    const topicTitle = document.getElementById('topicTitle');
    const topicDifficulty = document.getElementById('topicDifficulty');
    const topicDuration = document.getElementById('topicDuration');
    const topicExercises = document.getElementById('topicExercises');

    if (topicTitle) topicTitle.textContent = this.currentTopic.title;
    if (topicDifficulty) {
      topicDifficulty.textContent = this.currentTopic.difficulty <= 2 ? 'Beginner' : 
                                   this.currentTopic.difficulty <= 4 ? 'Intermediate' : 'Advanced';
      topicDifficulty.setAttribute('data-difficulty', this.currentTopic.difficulty);
    }
    if (topicDuration) topicDuration.textContent = this.currentTopic.duration;
    if (topicExercises) topicExercises.textContent = `${this.currentTopic.exercises} exercises`;

    // Ensure theory tab is active
    this.switchTab('theory');
  }

  showWelcomeScreen() {
    console.log('Showing welcome screen');
    const welcomeScreen = document.getElementById('welcomeScreen');
    const topicDetailView = document.getElementById('topicDetailView');

    if (welcomeScreen) {
      welcomeScreen.classList.remove('hidden');
    }
    if (topicDetailView) {
      topicDetailView.classList.add('hidden');
    }
    
    this.currentTopic = null;
    this.showNotification('Returned to topic selection');
  }

  renderTopicContent() {
    if (!this.currentTopic) return;

    console.log('Rendering topic content for:', this.currentTopic.title);
    this.renderLessonContent();
    this.renderPracticeExercises();
    this.generateQuiz();
  }

  // Lesson Management - Fixed version
  renderLessonContent() {
    console.log('Rendering lesson content');
    const lessonTitle = document.getElementById('lessonTitle');
    const lessonText = document.getElementById('lessonText');
    const lessonProgress = document.getElementById('lessonProgress');
    const lessonProgressText = document.getElementById('lessonProgressText');
    const prevBtn = document.getElementById('prevLessonBtn');
    const nextBtn = document.getElementById('nextLessonBtn');

    const lessons = this.getLessonsForTopic(this.currentTopic.id);
    const currentLessonData = lessons[this.currentLesson];

    if (lessonTitle) lessonTitle.textContent = currentLessonData.title;
    if (lessonText) lessonText.innerHTML = currentLessonData.content;
    
    const progressPercent = ((this.currentLesson + 1) / lessons.length) * 100;
    if (lessonProgress) lessonProgress.style.width = `${progressPercent}%`;
    if (lessonProgressText) lessonProgressText.textContent = `Lesson ${this.currentLesson + 1} of ${lessons.length}`;

    if (prevBtn) prevBtn.disabled = this.currentLesson === 0;
    if (nextBtn) nextBtn.disabled = this.currentLesson === lessons.length - 1;

    console.log('Lesson content rendered');
  }

  getLessonsForTopic(topicId) {
    const lessonContent = {
      'B1': [
        {
          title: 'Welcome to Excel',
          content: `
            <h4>Understanding the Excel Interface</h4>
            <p>Excel is a powerful spreadsheet application that helps you organize, analyze, and visualize data.</p>
            <ul>
              <li><strong>Workbook:</strong> The entire Excel file containing one or more worksheets</li>
              <li><strong>Worksheet:</strong> Individual tabs within a workbook, containing cells in a grid</li>
              <li><strong>Cell:</strong> The intersection of a row and column (e.g., A1, B2)</li>
              <li><strong>Ribbon:</strong> The toolbar at the top containing all Excel commands</li>
            </ul>
            <p><strong>Try this:</strong> Click on different cells in the practice spreadsheet to see how cell references change.</p>
          `
        },
        {
          title: 'Navigating Your Worksheet',
          content: `
            <h4>Moving Around Excel</h4>
            <p>Learn the fastest ways to navigate your spreadsheet:</p>
            <ul>
              <li><strong>Arrow Keys:</strong> Move one cell at a time</li>
              <li><strong>Ctrl + Arrow:</strong> Jump to the edge of data regions</li>
              <li><strong>Ctrl + Home:</strong> Go to cell A1</li>
              <li><strong>Ctrl + End:</strong> Go to the last used cell</li>
            </ul>
            <p><strong>Practice:</strong> Try clicking on cell A1, then B5, then C10 in the spreadsheet.</p>
          `
        },
        {
          title: 'The Ribbon and Tabs',
          content: `
            <h4>Understanding Excel's Command Structure</h4>
            <p>The Ribbon organizes Excel's features into logical groups:</p>
            <ul>
              <li><strong>Home:</strong> Basic formatting and editing tools</li>
              <li><strong>Insert:</strong> Charts, tables, and other objects</li>
              <li><strong>Page Layout:</strong> Print and page setup options</li>
              <li><strong>Formulas:</strong> Function library and formula tools</li>
              <li><strong>Data:</strong> Sorting, filtering, and data analysis</li>
            </ul>
            <p>Each tab contains related commands grouped into sections.</p>
          `
        }
      ]
    };

    // Return lessons for the topic, or default lessons if not found
    return lessonContent[topicId] || [
      {
        title: 'Lesson Content',
        content: `
          <h4>Welcome to ${this.currentTopic.title}</h4>
          <p>This topic covers: ${this.currentTopic.keyPoints.join(', ')}</p>
          <p>You'll complete ${this.currentTopic.exercises} exercises in approximately ${this.currentTopic.duration}.</p>
          <p><strong>Key Learning Points:</strong></p>
          <ul>${this.currentTopic.keyPoints.map(point => `<li>${point}</li>`).join('')}</ul>
          <p>Switch to the Practice tab to start working with the spreadsheet, or take the Quiz when you're ready to test your knowledge.</p>
        `
      }
    ];
  }

  nextLesson() {
    console.log('Next lesson clicked');
    const lessons = this.getLessonsForTopic(this.currentTopic.id);
    if (this.currentLesson < lessons.length - 1) {
      this.currentLesson++;
      this.renderLessonContent();
      this.showNotification(`Advanced to lesson ${this.currentLesson + 1}`);
    } else {
      this.showNotification('This is the last lesson. Try the Practice or Quiz tabs!');
    }
  }

  prevLesson() {
    if (this.currentLesson > 0) {
      this.currentLesson--;
      this.renderLessonContent();
      this.showNotification(`Returned to lesson ${this.currentLesson + 1}`);
    }
  }

  // Practice Exercises
  renderPracticeExercises() {
    const exercisesList = document.getElementById('exercisesList');
    const practiceTips = document.getElementById('practiceTips');
    
    if (!exercisesList || !this.currentTopic) return;

    const exercises = this.getExercisesForTopic(this.currentTopic.id);
    
    exercisesList.innerHTML = exercises.map((exercise, index) => `
      <div class="exercise-item" data-exercise="${index}">
        <div class="exercise-title">${exercise.title}</div>
        <div class="exercise-description">${exercise.description}</div>
      </div>
    `).join('');

    if (practiceTips) {
      const tips = this.getTipsForTopic(this.currentTopic.id);
      practiceTips.innerHTML = tips.map(tip => `<li>${tip}</li>`).join('');
    }
  }

  getExercisesForTopic(topicId) {
    const exercises = {
      'B1': [
        { title: 'Navigate to Cell D5', description: 'Click on cell D5 and observe the cell reference' },
        { title: 'Select Range A1:C3', description: 'Click and drag to select a 3x3 range of cells' },
        { title: 'Use Keyboard Navigation', description: 'Press Ctrl+End to go to the last used cell' }
      ]
    };

    return exercises[topicId] || [
      { title: 'Practice Exercise 1', description: 'Try entering data in the spreadsheet cells' },
      { title: 'Practice Exercise 2', description: 'Use the formula bar to enter a simple calculation' },
      { title: 'Practice Exercise 3', description: 'Load sample data and explore the spreadsheet features' }
    ];
  }

  getTipsForTopic(topicId) {
    const tips = {
      'B1': [
        'Use Ctrl+Arrow keys to quickly move between data regions',
        'The Name Box shows your current cell location',
        'Double-click sheet tabs to rename them'
      ]
    };

    return tips[topicId] || [
      'Practice regularly to build muscle memory',
      'Don\'t be afraid to experiment with different approaches',
      'Use the sample data to try real-world scenarios'
    ];
  }

  startExercise(exerciseIndex) {
    this.switchTab('practice');
    this.showNotification('Exercise started! Follow the instructions in the practice guide.');
  }

  // Tab Management - Fixed version
  switchTab(tabName) {
    console.log('Switching to tab:', tabName);

    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    const activeTab = document.querySelector(`[data-tab="${tabName}"]`);
    if (activeTab) {
      activeTab.classList.add('active');
    }

    // Update tab panes
    document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.add('hidden'));
    const activePane = document.getElementById(`${tabName}Pane`);
    if (activePane) {
      activePane.classList.remove('hidden');
    }

    // Initialize tab-specific content
    if (tabName === 'quiz') {
      this.initializeQuiz();
    } else if (tabName === 'practice') {
      // Ensure spreadsheet is rendered
      this.renderSpreadsheet();
    }

    console.log('Tab switched successfully');
  }

  // Quiz System
  generateQuiz() {
    if (!this.currentTopic) return;

    this.currentQuiz = this.getQuizForTopic(this.currentTopic.id);
    this.currentQuestionIndex = 0;
  }

  getQuizForTopic(topicId) {
    const quizzes = {
      'B1': {
        questions: [
          {
            question: 'What is the intersection of a row and column called?',
            options: ['Cell', 'Range', 'Workbook', 'Worksheet'],
            correct: 0
          },
          {
            question: 'Which key combination takes you to cell A1?',
            options: ['Ctrl+Home', 'Ctrl+End', 'Alt+Home', 'Shift+Home'],
            correct: 0
          },
          {
            question: 'What contains one or more worksheets?',
            options: ['Cell', 'Range', 'Workbook', 'Ribbon'],
            correct: 2
          }
        ]
      }
    };

    return quizzes[topicId] || {
      questions: [
        {
          question: `What is the main focus of the topic "${this.currentTopic.title}"?`,
          options: ['Data Analysis', 'Formatting', 'Calculations', 'All of the above'],
          correct: 3
        },
        {
          question: 'How many exercises does this topic contain?',
          options: [`${this.currentTopic.exercises-1}`, `${this.currentTopic.exercises}`, `${this.currentTopic.exercises+1}`, `${this.currentTopic.exercises+2}`],
          correct: 1
        }
      ]
    };
  }

  initializeQuiz() {
    if (!this.currentQuiz) return;
    
    this.renderCurrentQuestion();
    this.updateQuizProgress();
  }

  renderCurrentQuestion() {
    if (!this.currentQuiz) return;

    const question = this.currentQuiz.questions[this.currentQuestionIndex];
    const questionText = document.getElementById('questionText');
    const answerOptions = document.getElementById('answerOptions');
    const prevBtn = document.getElementById('prevQuestionBtn');
    const nextBtn = document.getElementById('nextQuestionBtn');
    const submitBtn = document.getElementById('submitQuizBtn');

    if (questionText) questionText.textContent = question.question;
    
    if (answerOptions) {
      answerOptions.innerHTML = question.options.map((option, index) => `
        <div class="answer-option" data-option="${index}">
          ${String.fromCharCode(65 + index)}. ${option}
        </div>
      `).join('');
    }

    if (prevBtn) prevBtn.disabled = this.currentQuestionIndex === 0;
    
    const isLastQuestion = this.currentQuestionIndex === this.currentQuiz.questions.length - 1;
    if (nextBtn) nextBtn.style.display = isLastQuestion ? 'none' : 'block';
    if (submitBtn) submitBtn.style.display = isLastQuestion ? 'block' : 'none';
  }

  updateQuizProgress() {
    if (!this.currentQuiz) return;

    const progress = ((this.currentQuestionIndex + 1) / this.currentQuiz.questions.length) * 100;
    const progressBar = document.getElementById('quizProgress');
    const progressText = document.getElementById('quizProgressText');

    if (progressBar) progressBar.style.width = `${progress}%`;
    if (progressText) {
      progressText.textContent = `Question ${this.currentQuestionIndex + 1} of ${this.currentQuiz.questions.length}`;
    }
  }

  selectAnswer(element) {
    // Remove previous selection
    document.querySelectorAll('.answer-option').forEach(opt => opt.classList.remove('selected'));
    
    // Add selection to clicked option
    element.classList.add('selected');
    
    // Store the answer
    if (!this.currentQuiz.userAnswers) {
      this.currentQuiz.userAnswers = [];
    }
    this.currentQuiz.userAnswers[this.currentQuestionIndex] = parseInt(element.getAttribute('data-option'));
  }

  nextQuestion() {
    if (this.currentQuestionIndex < this.currentQuiz.questions.length - 1) {
      this.currentQuestionIndex++;
      this.renderCurrentQuestion();
      this.updateQuizProgress();
    }
  }

  prevQuestion() {
    if (this.currentQuestionIndex > 0) {
      this.currentQuestionIndex--;
      this.renderCurrentQuestion();
      this.updateQuizProgress();
    }
  }

  submitQuiz() {
    if (!this.currentQuiz || !this.currentQuiz.userAnswers) {
      this.showNotification('Please answer all questions before submitting.', 'warning');
      return;
    }

    // Calculate score
    let correctAnswers = 0;
    this.currentQuiz.questions.forEach((question, index) => {
      if (this.currentQuiz.userAnswers[index] === question.correct) {
        correctAnswers++;
      }
    });

    const score = Math.round((correctAnswers / this.currentQuiz.questions.length) * 100);
    
    // Store the score
    this.userProgress.quizScores[this.currentTopic.id] = score;
    
    // Show results
    this.showQuizResults(score, correctAnswers);
    
    // Check for achievements
    if (score === 100) {
      this.earnAchievement('perfect');
    }
  }

  showQuizResults(score, correctAnswers) {
    const quizContent = document.getElementById('quizContent');
    const quizResults = document.getElementById('quizResults');

    if (quizContent) quizContent.classList.add('hidden');
    if (quizResults) quizResults.classList.remove('hidden');

    const scorePercentage = document.getElementById('quizScore');
    const scoreMessage = document.getElementById('scoreMessage');

    if (scorePercentage) scorePercentage.textContent = `${score}%`;
    
    let message = '';
    if (score >= 90) message = 'Excellent work! You\'ve mastered this topic.';
    else if (score >= 75) message = 'Great job! You have a solid understanding.';
    else if (score >= 60) message = 'Good effort! Review the lessons and try again.';
    else message = 'Keep practicing! Review the material and retake the quiz.';
    
    if (scoreMessage) scoreMessage.textContent = message;

    // Update progress circle
    const scoreCircle = document.querySelector('.score-circle');
    if (scoreCircle) {
      scoreCircle.style.setProperty('--score-angle', `${(score / 100) * 360}deg`);
    }
  }

  completeTopic() {
    if (!this.currentTopic) return;

    // Mark topic as completed
    if (!this.userProgress.completedTopics.includes(this.currentTopic.id)) {
      this.userProgress.completedTopics.push(this.currentTopic.id);
    }

    // Check for achievements
    this.checkTopicAchievements();
    
    // Update progress
    this.updateProgress();
    
    // Show completion message
    this.showNotification(`🎉 Congratulations! You've completed "${this.currentTopic.title}"`, 'success');
    
    // Return to topics view
    setTimeout(() => {
      this.showWelcomeScreen();
    }, 2000);
  }

  checkTopicAchievements() {
    const topicId = this.currentTopic.id;
    
    // First beginner topic
    if (topicId.startsWith('B') && this.userProgress.completedTopics.filter(id => id.startsWith('B')).length === 1) {
      this.earnAchievement('beginner');
    }
    
    // All basic formulas
    if (topicId === 'B3') {
      this.earnAchievement('formulas');
    }
    
    // Charts topic
    if (topicId === 'B6' || topicId === 'I7') {
      this.earnAchievement('charts');
    }
    
    // Pivot tables
    if (topicId === 'I8' || topicId === 'A1') {
      this.earnAchievement('pivot');
    }
    
    // Advanced topic
    if (topicId.startsWith('A')) {
      this.earnAchievement('advanced');
    }
  }

  // Spreadsheet Management - Fixed version
  renderSpreadsheet() {
    const header = document.getElementById('spreadsheetHeader');
    const body = document.getElementById('spreadsheetBody');
    
    if (!header || !body) {
      console.log('Spreadsheet elements not found');
      return;
    }

    // Create header row
    header.innerHTML = '<tr><th></th>' + 
      Array.from({length: 15}, (_, i) => `<th>${String.fromCharCode(65 + i)}</th>`).join('') + '</tr>';

    // Create data rows
    body.innerHTML = '';
    for (let row = 0; row < 30; row++) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<th>${row + 1}</th>` +
        Array.from({length: 15}, (_, col) => {
          const value = this.spreadsheetData[row][col];
          return `<td data-row="${row}" data-col="${col}">
            <span class="cell-display">${this.formatCellValue(value)}</span>
          </td>`;
        }).join('');
      body.appendChild(tr);
    }

    console.log('Spreadsheet rendered');
  }

  renderPracticeSpreadsheet() {
    const header = document.getElementById('practiceSpreadsheetHeader');
    const body = document.getElementById('practiceSpreadsheetBody');
    
    if (!header || !body) return;

    // Create header row
    header.innerHTML = '<tr><th></th>' + 
      Array.from({length: 15}, (_, i) => `<th>${String.fromCharCode(65 + i)}</th>`).join('') + '</tr>';

    // Create data rows
    body.innerHTML = '';
    for (let row = 0; row < 30; row++) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<th>${row + 1}</th>` +
        Array.from({length: 15}, (_, col) => {
          const value = this.practiceSpreadsheetData[row][col];
          return `<td data-practice-row="${row}" data-practice-col="${col}">
            <span class="cell-display">${this.formatCellValue(value)}</span>
          </td>`;
        }).join('');
      body.appendChild(tr);
    }
  }

  formatCellValue(value) {
    if (value === null || value === undefined || value === '') return '';
    
    // Handle formulas
    if (typeof value === 'string' && value.startsWith('=')) {
      try {
        const result = this.evaluateFormula(value);
        return result.toString();
      } catch (e) {
        return '#ERROR!';
      }
    }
    
    return value.toString();
  }

  selectCell(row, col) {
    console.log(`Selecting cell ${String.fromCharCode(65 + col)}${row + 1}`);
    
    // Remove previous selection
    document.querySelectorAll('#spreadsheet td.selected').forEach(cell => {
      cell.classList.remove('selected');
    });
    
    // Add selection to new cell
    const cell = document.querySelector(`#spreadsheet td[data-row="${row}"][data-col="${col}"]`);
    if (cell) {
      cell.classList.add('selected');
      this.selectedCell = { row, col };
      
      // Update formula bar
      const formulaInput = document.getElementById('formulaInput');
      if (formulaInput) {
        formulaInput.value = this.spreadsheetData[row][col] || '';
      }

      this.showNotification(`Selected cell ${String.fromCharCode(65 + col)}${row + 1}`);
    }
  }

  selectPracticeCell(row, col) {
    // Remove previous selection
    document.querySelectorAll('#practiceSpreadsheet td.selected').forEach(cell => {
      cell.classList.remove('selected');
    });
    
    // Add selection to new cell
    const cell = document.querySelector(`#practiceSpreadsheet td[data-practice-row="${row}"][data-practice-col="${col}"]`);
    if (cell) {
      cell.classList.add('selected');
      this.selectedPracticeCell = { row, col };
      
      // Update formula bar
      const formulaInput = document.getElementById('practiceFormulaInput');
      if (formulaInput) {
        formulaInput.value = this.practiceSpreadsheetData[row][col] || '';
      }
    }
  }

  executeFormula() {
    const formulaInput = document.getElementById('formulaInput');
    if (!formulaInput) return;

    const value = formulaInput.value;
    this.updateCell(this.selectedCell.row, this.selectedCell.col, value);
  }

  executePracticeFormula() {
    const formulaInput = document.getElementById('practiceFormulaInput');
    if (!formulaInput) return;

    const value = formulaInput.value;
    this.updatePracticeCell(this.selectedPracticeCell.row, this.selectedPracticeCell.col, value);
  }

  updateCell(row, col, value) {
    // Save state for undo
    this.undoStack.push(JSON.parse(JSON.stringify(this.spreadsheetData)));
    this.redoStack = [];

    // Update data
    this.spreadsheetData[row][col] = value;

    // Re-render spreadsheet
    this.renderSpreadsheet();

    // Select the updated cell
    this.selectCell(row, col);

    const cellRef = String.fromCharCode(65 + col) + (row + 1);
    this.showNotification(`Updated cell ${cellRef}`);
  }

  updatePracticeCell(row, col, value) {
    // Save state for undo
    this.practiceUndoStack.push(JSON.parse(JSON.stringify(this.practiceSpreadsheetData)));
    this.practiceRedoStack = [];

    // Update data
    this.practiceSpreadsheetData[row][col] = value;

    // Re-render spreadsheet
    this.renderPracticeSpreadsheet();

    // Select the updated cell
    this.selectPracticeCell(row, col);

    const cellRef = String.fromCharCode(65 + col) + (row + 1);
    this.showNotification(`Updated practice cell ${cellRef}`);
  }

  // Formula Evaluation
  evaluateFormula(formula) {
    if (!formula.startsWith('=')) return formula;

    let expression = formula.substring(1);

    // Handle functions
    expression = this.replaceFunctions(expression);

    // Handle cell references
    expression = expression.replace(/[A-Z]+\d+/g, (match) => {
      return this.getCellValue(match).toString();
    });

    // Evaluate safely
    try {
      if (/^[\d\s+\-*/.()]+$/.test(expression)) {
        return eval(expression);
      } else {
        return '#ERROR!';
      }
    } catch (e) {
      return '#ERROR!';
    }
  }

  replaceFunctions(expression) {
    // SUM function
    expression = expression.replace(/SUM\(([^)]+)\)/g, (match, args) => {
      const values = this.parseRange(args);
      return values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
    });

    // AVERAGE function
    expression = expression.replace(/AVERAGE\(([^)]+)\)/g, (match, args) => {
      const values = this.parseRange(args);
      const numericValues = values.filter(val => !isNaN(parseFloat(val)) && val !== '');
      return numericValues.length ? 
        numericValues.reduce((sum, val) => sum + parseFloat(val), 0) / numericValues.length : 0;
    });

    // COUNT function
    expression = expression.replace(/COUNT\(([^)]+)\)/g, (match, args) => {
      const values = this.parseRange(args);
      return values.filter(val => !isNaN(parseFloat(val)) && val !== '').length;
    });

    // TODAY function
    expression = expression.replace(/TODAY\(\)/g, () => {
      return '"' + new Date().toLocaleDateString() + '"';
    });

    return expression;
  }

  parseRange(rangeStr) {
    const values = [];
    const parts = rangeStr.split(',');

    parts.forEach(part => {
      part = part.trim();
      if (part.includes(':')) {
        // Range like A1:A5
        const [start, end] = part.split(':');
        const startCol = start.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
        const startRow = parseInt(start.match(/\d+/)[0]) - 1;
        const endCol = end.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
        const endRow = parseInt(end.match(/\d+/)[0]) - 1;

        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            if (row < this.spreadsheetData.length && col < this.spreadsheetData[0].length) {
              values.push(this.spreadsheetData[row][col] || '');
            }
          }
        }
      } else {
        // Single cell
        values.push(this.getCellValue(part));
      }
    });

    return values;
  }

  getCellValue(cellRef) {
    const col = cellRef.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
    const row = parseInt(cellRef.match(/\d+/)[0]) - 1;

    if (row >= 0 && row < this.spreadsheetData.length && 
        col >= 0 && col < this.spreadsheetData[0].length) {
      let value = this.spreadsheetData[row][col];
      if (typeof value === 'string' && value.startsWith('=')) {
        return this.evaluateFormula(value);
      }
      return value || 0;
    }
    return 0;
  }

  // Data Management - Fixed version
  loadSampleData(datasetName) {
    console.log('Loading sample data:', datasetName);
    const dataset = appData.sampleDatasets[datasetName];
    if (!dataset) {
      console.error('Dataset not found:', datasetName);
      return;
    }

    // Clear existing data
    this.spreadsheetData = Array(30).fill().map(() => Array(15).fill(''));

    // Load sample data
    dataset.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (rowIndex < this.spreadsheetData.length && colIndex < this.spreadsheetData[0].length) {
          this.spreadsheetData[rowIndex][colIndex] = cell;
        }
      });
    });

    this.renderSpreadsheet();
    this.showNotification(`Loaded ${datasetName} dataset successfully!`);
  }

  loadPracticeData() {
    const select = document.getElementById('practiceDataset');
    if (!select || !select.value) return;

    const dataset = appData.sampleDatasets[select.value];
    if (!dataset) return;

    // Clear existing data
    this.practiceSpreadsheetData = Array(30).fill().map(() => Array(15).fill(''));

    // Load sample data
    dataset.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (rowIndex < this.practiceSpreadsheetData.length && colIndex < this.practiceSpreadsheetData[0].length) {
          this.practiceSpreadsheetData[rowIndex][colIndex] = cell;
        }
      });
    });

    this.renderPracticeSpreadsheet();
    this.showNotification(`Loaded ${select.value} data into practice mode!`);
  }

  // Undo/Redo
  undo() {
    if (this.undoStack.length > 0) {
      this.redoStack.push(JSON.parse(JSON.stringify(this.spreadsheetData)));
      this.spreadsheetData = this.undoStack.pop();
      this.renderSpreadsheet();
      this.showNotification('Undid last action');
    }
  }

  redo() {
    if (this.redoStack.length > 0) {
      this.undoStack.push(JSON.parse(JSON.stringify(this.spreadsheetData)));
      this.spreadsheetData = this.redoStack.pop();
      this.renderSpreadsheet();
      this.showNotification('Redid last action');
    }
  }

  clearSpreadsheet() {
    this.undoStack.push(JSON.parse(JSON.stringify(this.spreadsheetData)));
    this.spreadsheetData = Array(30).fill().map(() => Array(15).fill(''));
    this.renderSpreadsheet();
    this.showNotification('Spreadsheet cleared');
  }

  save() {
    // In a real app, this would save to a server or local storage
    this.showNotification('Spreadsheet saved successfully!');
  }

  // Modal Management - Fixed version
  showModal(modalName) {
    console.log('Showing modal:', modalName);
    const modal = document.getElementById(`${modalName}Modal`);
    if (modal) {
      modal.classList.remove('hidden');
      console.log('Modal shown:', modalName);
      
      // Initialize modal content if needed
      if (modalName === 'practiceMode') {
        this.renderPracticeSpreadsheet();
      }
    } else {
      console.error('Modal not found:', modalName);
    }
  }

  hideModal(modalName) {
    console.log('Hiding modal:', modalName);
    const modal = document.getElementById(`${modalName}Modal`);
    if (modal) {
      modal.classList.add('hidden');
    }
  }

  // Chat System
  sendChatMessage() {
    const chatInput = document.getElementById('chatInput');
    const chatContainer = document.getElementById('chatContainer');
    
    if (!chatInput || !chatContainer || !chatInput.value.trim()) return;

    const message = chatInput.value.trim();

    // Add user message
    const userMessage = document.createElement('div');
    userMessage.className = 'chat-message user';
    userMessage.innerHTML = `
      <div class="message-avatar">👤</div>
      <div class="message-content">${message}</div>
    `;
    chatContainer.appendChild(userMessage);

    // Generate AI response
    setTimeout(() => {
      const response = this.generateAIResponse(message);
      const botMessage = document.createElement('div');
      botMessage.className = 'chat-message bot';
      botMessage.innerHTML = `
        <div class="message-avatar">🤖</div>
        <div class="message-content">${response}</div>
      `;
      chatContainer.appendChild(botMessage);
      chatContainer.scrollTop = chatContainer.scrollHeight;
    }, 1000);

    chatInput.value = '';
    chatContainer.scrollTop = chatContainer.scrollHeight;
  }

  generateAIResponse(message) {
    const lowerMessage = message.toLowerCase();
    
    const responses = {
      'formula': 'Excel formulas always start with = (equals sign). Try =SUM(A1:A5) to add values in cells A1 through A5. What specific formula would you like help with?',
      'sum': 'The SUM function adds numbers together. Use =SUM(range) like =SUM(A1:A10). You can also add individual cells: =SUM(A1,C1,E1).',
      'vlookup': 'VLOOKUP searches for a value in the first column and returns data from another column. Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, FALSE). The FALSE ensures exact match.',
      'if': 'IF statements test conditions. Syntax: =IF(condition, value_if_true, value_if_false). Example: =IF(A1>10,"High","Low").',
      'chart': 'To create charts: 1) Select your data, 2) Use the Charts tool in the sidebar, 3) Choose chart type. Bar charts compare categories, line charts show trends.',
      'pivot': 'Pivot tables summarize data. Select your data range, then create a pivot table to group and analyze information quickly.',
      'error': 'Common Excel errors: #DIV/0! (division by zero), #VALUE! (wrong data type), #REF! (invalid cell reference), #NAME? (unrecognized function).',
      'help': 'I can help with Excel formulas, functions, charts, pivot tables, cell references, and more! What specific topic interests you?'
    };

    // Find matching response
    for (const [key, response] of Object.entries(responses)) {
      if (lowerMessage.includes(key)) {
        return response;
      }
    }

    // Default responses based on context
    if (this.currentTopic) {
      return `Great question about "${this.currentTopic.title}"! Try practicing with the exercises in the Practice tab. You can also review the lesson content for step-by-step guidance.`;
    }

    return 'I\'m here to help with Excel! You can ask me about formulas (SUM, VLOOKUP, IF), creating charts, pivot tables, or any Excel concept. What would you like to learn?';
  }

  // Chart Creation
  createChart() {
    const chartType = document.getElementById('chartType').value;
    const canvas = document.getElementById('chart');
    
    if (!canvas) return;

    // Destroy existing chart
    if (this.chart) {
      this.chart.destroy();
    }

    // Use sample data
    const chartData = {
      labels: ['North', 'South', 'East', 'West'],
      datasets: [{
        label: 'Sales Data',
        data: [15000, 12000, 18000, 14000],
        backgroundColor: ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5'],
        borderColor: ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5'],
        borderWidth: 1
      }]
    };

    const config = {
      type: chartType,
      data: chartData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: 'Excel Data Visualization',
            color: '#f5f5f5'
          },
          legend: {
            labels: {
              color: '#f5f5f5'
            }
          }
        },
        scales: chartType !== 'pie' && chartType !== 'doughnut' ? {
          y: {
            beginAtZero: true,
            ticks: {
              color: '#f5f5f5'
            },
            grid: {
              color: 'rgba(245, 245, 245, 0.1)'
            }
          },
          x: {
            ticks: {
              color: '#f5f5f5'
            },
            grid: {
              color: 'rgba(245, 245, 245, 0.1)'
            }
          }
        } : {}
      }
    };

    this.chart = new Chart(canvas, config);
    this.showNotification(`${chartType.charAt(0).toUpperCase() + chartType.slice(1)} chart created!`);
  }

  // Functions Reference
  renderFunctions() {
    const functionsGrid = document.getElementById('functionsGrid');
    if (!functionsGrid) return;

    const functions = Object.entries(excelFunctions);
    this.renderFunctionsList(functions);
  }

  renderFunctionsList(functions) {
    const functionsGrid = document.getElementById('functionsGrid');
    if (!functionsGrid) return;

    functionsGrid.innerHTML = functions.map(([name, info]) => `
      <div class="function-item" data-function="${name}">
        <div class="function-name">${name}</div>
        <div class="function-description">${info.description}</div>
        <div class="function-syntax">${info.syntax}</div>
      </div>
    `).join('');
  }

  filterFunctions(searchTerm) {
    const functions = Object.entries(excelFunctions);
    const filtered = functions.filter(([name, info]) =>
      name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      info.description.toLowerCase().includes(searchTerm.toLowerCase())
    );
    this.renderFunctionsList(filtered);
  }

  insertFunction(functionName) {
    const formulaInput = document.getElementById('formulaInput');
    const practiceFormulaInput = document.getElementById('practiceFormulaInput');
    
    const syntax = excelFunctions[functionName]?.syntax || `=${functionName}()`;
    
    if (formulaInput && !formulaInput.closest('.modal.hidden')) {
      formulaInput.value = syntax;
      formulaInput.focus();
    } else if (practiceFormulaInput && !practiceFormulaInput.closest('.modal.hidden')) {
      practiceFormulaInput.value = syntax;
      practiceFormulaInput.focus();
    }
    
    this.hideModal('functions');
    this.showNotification(`Inserted ${functionName} function`);
  }

  // Achievements System
  renderAchievements() {
    const achievementsGrid = document.getElementById('achievementsGrid');
    if (!achievementsGrid) return;

    achievementsGrid.innerHTML = appData.achievements.map(achievement => `
      <div class="achievement-item ${this.userProgress.earnedAchievements.includes(achievement.id) ? 'earned' : ''}">
        <div class="achievement-icon">${achievement.icon}</div>
        <div class="achievement-name">${achievement.name}</div>
        <div class="achievement-description">${achievement.description}</div>
      </div>
    `).join('');
  }

  earnAchievement(achievementId) {
    if (this.userProgress.earnedAchievements.includes(achievementId)) return;

    this.userProgress.earnedAchievements.push(achievementId);
    const achievement = appData.achievements.find(a => a.id === achievementId);
    
    if (achievement) {
      this.showNotification(`🏆 Achievement Unlocked: ${achievement.name}!`, 'success');
      this.renderAchievements();
      this.updateProgress();
    }
  }

  // Progress Management
  updateProgress() {
    const totalTopics = appData.topicCategories.reduce((sum, cat) => sum + cat.topics.length, 0);
    const completedCount = this.userProgress.completedTopics.length;
    const progressPercentage = Math.round((completedCount / totalTopics) * 100);

    this.userProgress.totalProgress = progressPercentage;

    const progressFill = document.getElementById('overallProgress');
    const progressText = document.getElementById('progressText');

    if (progressFill) progressFill.style.width = `${progressPercentage}%`;
    if (progressText) progressText.textContent = `${progressPercentage}% Complete`;

    // Update topic counts
    this.renderTopics();
  }

  // Notification System
  showNotification(message, type = 'success') {
    console.log('Showing notification:', message, type);
    const notification = document.getElementById('notification');
    const notificationText = document.getElementById('notificationText');
    const notificationIcon = document.getElementById('notificationIcon');

    if (!notification || !notificationText || !notificationIcon) return;

    // Set icon based on type
    const icons = {
      success: '✅',
      error: '❌',
      warning: '⚠️',
      info: 'ℹ️'
    };

    notificationIcon.textContent = icons[type] || icons.success;
    notificationText.textContent = message;

    // Update notification style
    notification.className = `notification ${type}`;
    notification.classList.remove('hidden');

    // Auto-hide after 4 seconds
    setTimeout(() => {
      this.hideNotification();
    }, 4000);
  }

  hideNotification() {
    const notification = document.getElementById('notification');
    if (notification) {
      notification.classList.add('hidden');
    }
  }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
  console.log('DOM loaded, initializing app...');
  window.excelApp = new ExcelLearningApp();
});