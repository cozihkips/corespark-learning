<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CoreSpark - School Exam Analytics</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: #f0f4f8; color: #333; }

        /* ── LOGIN SCREEN ── */
        #loginScreen {
            position: fixed; inset: 0; z-index: 9999;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #1e40af 100%);
            display: flex; align-items: center; justify-content: center;
            flex-direction: column; min-height: 100vh;
        }
        #loginScreen.hidden { display: none; }
        .login-box {
            background: white; border-radius: 10px; padding: 36px 36px 28px;
            width: 100%; max-width: 380px; box-shadow: 0 20px 60px rgba(0,0,0,0.4);
        }
        .login-logo {
            text-align: center; margin-bottom: 24px;
        }
        .login-logo .spark { font-size: 38px; }
        .login-logo h2 { font-size: 22px; font-weight: 900; color: #1e40af; letter-spacing: 0.5px; margin-top: 6px; }
        .login-logo p { font-size: 12px; color: #6b7280; margin-top: 2px; }
        .login-field { margin-bottom: 16px; }
        .login-field label { display: block; font-size: 12px; font-weight: 700; color: #374151; margin-bottom: 5px; text-transform: uppercase; letter-spacing: 0.4px; }
        .login-field input {
            width: 100%; padding: 11px 14px; border: 1.5px solid #d1d5db; border-radius: 5px;
            font-size: 14px; transition: border-color 0.2s;
        }
        .login-field input:focus { outline: none; border-color: #1e40af; box-shadow: 0 0 0 3px rgba(30,64,175,0.12); }
        .login-btn {
            width: 100%; padding: 12px; background: #1e40af; color: white; border: none;
            border-radius: 5px; font-size: 15px; font-weight: 800; cursor: pointer;
            letter-spacing: 0.3px; margin-top: 4px; transition: background 0.2s;
        }
        .login-btn:hover { background: #1e3a8a; }
        .login-error { background: #fef2f2; color: #991b1b; border: 1px solid #fca5a5; border-radius: 4px; padding: 10px 14px; font-size: 13px; margin-bottom: 14px; display: none; }
        .login-error.show { display: block; }
        .login-footer { text-align: center; margin-top: 20px; font-size: 11px; color: #9ca3af; }

        /* ── APP (hidden until logged in) ── */
        #appRoot { display: none; }
        #appRoot.visible { display: block; }

        /* ── TOP BAR logout button ── */
        .logout-btn {
            background: rgba(255,255,255,0.18); color: white; border: 1px solid rgba(255,255,255,0.35);
            padding: 5px 13px; border-radius: 4px; font-size: 12px; font-weight: 700; cursor: pointer;
        }
        .logout-btn:hover { background: rgba(255,255,255,0.3); }
        .header { background: #1e40af; color: white; padding: 16px 20px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px; }
        .header h1 { font-size: 22px; }
        .school-name { font-size: 13px; opacity: 0.85; }
        .container { max-width: 1150px; margin: 0 auto; padding: 16px; }
        .section { background: white; padding: 20px; margin-bottom: 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
        .section-title { font-size: 16px; font-weight: bold; margin-bottom: 14px; color: #1e3a8a; border-bottom: 2px solid #dbeafe; padding-bottom: 8px; }
        .tabs { display: flex; flex-wrap: wrap; background: white; border-radius: 6px 6px 0 0; box-shadow: 0 1px 3px rgba(0,0,0,0.08); padding: 8px 8px 0; margin-bottom: 0; }
        .tab { padding: 8px 11px; border: none; background: none; cursor: pointer; color: #555; font-weight: bold; font-size: 12px; border-bottom: 3px solid transparent; }
        .tab.active { color: #1e40af; border-bottom: 3px solid #1e40af; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .input-group { margin-bottom: 12px; }
        .input-group label { display: block; margin-bottom: 4px; font-weight: bold; font-size: 13px; }
        .input-group input, .input-group select { width: 100%; padding: 9px 10px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 14px; }
        .input-group input:focus, .input-group select:focus { outline: none; border-color: #3b82f6; }
        .btn { padding: 9px 16px; background: #1e40af; color: white; border: none; cursor: pointer; border-radius: 4px; font-size: 13px; font-weight: bold; margin-bottom: 8px; margin-right: 6px; }
        .btn:hover { background: #1e3a8a; }
        .btn-green { background: #15803d; } .btn-green:hover { background: #166534; }
        .btn-red { background: #dc2626; } .btn-red:hover { background: #b91c1c; }
        .tbl-wrap { overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 500px; }
        th, td { padding: 8px 10px; text-align: left; border-bottom: 1px solid #e5e7eb; white-space: nowrap; }
        th { background: #eff6ff; font-weight: bold; color: #1e3a8a; font-size: 12px; }
        tr:hover td { background: #f9fafb; }
        .message { padding: 11px 14px; margin-bottom: 14px; border-radius: 4px; display: none; font-size: 13px; }
        .message.show { display: block; }
        .message.success { background: #dcfce7; color: #14532d; border-left: 4px solid #16a34a; }
        .message.error { background: #fef2f2; color: #7f1d1d; border-left: 4px solid #dc2626; }
        .message.info { background: #eff6ff; color: #1e3a8a; border-left: 4px solid #3b82f6; }
        .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; }
        .item { padding: 9px 12px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; margin-bottom: 6px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; }
        .item-delete { background: #fee2e2; color: #dc2626; border: 1px solid #fca5a5; padding: 3px 8px; font-size: 11px; cursor: pointer; border-radius: 3px; font-weight: bold; }
        .stat-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 10px; margin-bottom: 16px; }
        .stat-card { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 12px; text-align: center; }
        .stat-card .val { font-size: 20px; font-weight: bold; color: #1e40af; }
        .stat-card .lbl { font-size: 11px; color: #6b7280; margin-top: 3px; }
        .badge { display: inline-block; padding: 2px 7px; border-radius: 10px; font-size: 11px; font-weight: bold; }
        .badge-0 { background: #dcfce7; color: #14532d; }
        .badge-1 { background: #dbeafe; color: #1e3a8a; }
        .badge-2 { background: #fef9c3; color: #713f12; }
        .badge-3 { background: #fee2e2; color: #7f1d1d; }
        .badge-x { background: #f3f4f6; color: #374151; }
        .block { border: 1px solid #e2e8f0; border-radius: 6px; margin-bottom: 14px; overflow: hidden; }
        .block-hdr { background: #eff6ff; padding: 10px 14px; font-weight: bold; font-size: 13px; color: #1e3a8a; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px; }
        .block-sub { background: #f8fafc; padding: 8px 14px; font-size: 12px; color: #6b7280; border-bottom: 1px solid #e2e8f0; }
        .mode-tabs { display: flex; gap: 6px; margin-bottom: 16px; flex-wrap: wrap; }
        .mode-tab { padding: 7px 13px; border: 1px solid #bfdbfe; background: white; color: #1e40af; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: bold; }
        .mode-tab.active { background: #1e40af; color: white; }
        .info-box { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 4px; padding: 12px; margin-bottom: 14px; font-size: 13px; line-height: 1.6; }
        .pct-bar-wrap { background: #e5e7eb; border-radius: 3px; height: 5px; width: 60px; display: inline-block; vertical-align: middle; margin-left: 5px; }
        .pct-bar { background: #3b82f6; border-radius: 3px; height: 5px; }
        .stream-tag { display: inline-block; background: #dbeafe; color: #1e3a8a; border-radius: 3px; padding: 1px 6px; font-size: 11px; font-weight: bold; margin-right: 3px; }
        .grade-compare-hdr { background: #1e40af; color: white; padding: 8px 14px; font-weight: bold; font-size: 13px; border-radius: 4px 4px 0 0; margin-top: 16px; }
    </style>
</head>
<body>
<!-- ══ LOGIN SCREEN ══ -->
<div id="loginScreen">
    <div class="login-box">
        <div class="login-logo">
            <div class="spark">⚡</div>
            <h2>CoreSpark</h2>
            <p>School Exam Analytics System</p>
        </div>
        <div id="loginError" class="login-error"></div>
        <div class="login-field">
            <label>Username</label>
            <input type="text" id="loginUsername" placeholder="Enter username" autocomplete="username"
                onkeydown="if(event.key==='Enter') document.getElementById('loginPassword').focus()">
        </div>
        <div class="login-field">
            <label>Password</label>
            <input type="password" id="loginPassword" placeholder="Enter password" autocomplete="current-password"
                onkeydown="if(event.key==='Enter') doLogin()">
        </div>
        <button class="login-btn" onclick="doLogin()">Sign In</button>
        <div class="login-footer">CoreSpark · MM Chandaria Primary &amp; Junior School</div>
    </div>
</div>

<!-- ══ APP ROOT ══ -->
<div id="appRoot">
<div class="header">
    <h1>⚡ CoreSpark</h1>
    <div class="school-name" id="schoolDisplay">School: Not Set</div>
    <div style="display:flex;align-items:center;gap:10px;">
        <span id="loggedInUser" style="font-size:12px;opacity:0.8;"></span>
        <button class="logout-btn" onclick="doLogout()">Sign Out</button>
    </div>
</div>

<div class="container">
    <div id="message" class="message"></div>

    <!-- Teacher role banner -->
    <div id="teacherBanner" style="display:none;background:#fef9c3;border:1px solid #fde047;border-radius:6px;padding:10px 16px;margin-bottom:10px;font-size:13px;color:#713f12;">
        <strong>👩‍🏫 Teacher View:</strong> You can enter marks for active exams and view past results. Contact an admin to change settings.
    </div>

    <div class="tabs"></div>
        <button class="tab active" onclick="switchTab(0,this)">🏫 School</button>
        <button class="tab" onclick="switchTab(12,this)">👨‍🏫 Teachers</button>
        <button class="tab" onclick="switchTab(2,this)">📖 Subjects</button>
        <button class="tab" onclick="switchTab(3,this)">⚙️ Levels</button>
        <button class="tab" onclick="switchTab(13,this)">📝 Exams</button>
        <button class="tab" onclick="switchTab(14,this)">⚡ Import All-in-One</button>
        <button class="tab" onclick="switchTab(8,this)">👥 Import Learners</button>
        <button class="tab" onclick="switchTab(10,this)">🗂️ Learners Store</button>
        <button class="tab" onclick="switchTab(9,this)">📊 Import Marks</button>
        <button class="tab" onclick="switchTab(11,this)">📑 Marks Store</button>
        <button class="tab" onclick="switchTab(5,this)">📈 Analysis</button>
        <button class="tab" onclick="switchTab(6,this)">📋 Past Results</button>
        <button class="tab" onclick="switchTab(7,this)">📥 Downloads</button>
        <button class="tab" onclick="switchTab(1,this)">📚 Classes</button>
        <button class="tab" onclick="switchTab(4,this)">📤 Enter Marks</button>
    </div>

    <!-- TAB 0: SCHOOL -->
    <div id="tab-0" class="tab-content active">
        <div class="section">
            <div class="section-title">🏫 School Information</div>
            <div class="grid-2">
                <div>
                    <div class="input-group"><label>School Name</label><input type="text" id="schoolName" placeholder="Enter your school name"></div>
                    <div class="input-group"><label>School Motto / Tagline (optional)</label><input type="text" id="schoolMotto" placeholder="e.g. Excellence in Education"></div>
                    <div class="input-group"><label>School Address / Location (optional)</label><input type="text" id="schoolAddress" placeholder="e.g. Nairobi, Kenya"></div>
                    <button class="btn" onclick="saveSchoolData()">💾 Save School Information</button>
                </div>
                <div>
                    <div class="input-group"><label>School Logo</label>
                        <input type="file" id="logoFile" accept="image/*" onchange="uploadLogo()">
                    </div>
                    <div id="logoPreview" style="margin-top:10px;"></div>
                    <div id="logoStatus" style="font-size:12px;color:#15803d;margin-top:6px;"></div>
                </div>
            </div>
        </div>

        <!-- User Management — admin only -->
        <div class="section">
            <div class="section-title">👤 User Management</div>
            <p style="font-size:13px;color:#6b7280;margin-bottom:14px;">Manage who can access CoreSpark. Only admin users can add or remove accounts.</p>
            <div id="userManagerContent" style="margin-bottom:16px;"></div>
            <div class="block" id="addUserBlock">
                <div class="block-hdr">➕ Add New User</div>
                <div style="padding:14px;">
                    <div class="grid-2">
                        <div class="input-group"><label>Username</label>
                            <input type="text" id="newUsername" placeholder="e.g. jkamau">
                        </div>
                        <div class="input-group"><label>Password</label>
                            <input type="password" id="newPassword" placeholder="Min 4 characters">
                        </div>
                        <div class="input-group"><label>Role</label>
                            <select id="newUserRole">
                                <option value="teacher">Teacher</option>
                                <option value="admin">Admin</option>
                            </select>
                        </div>
                    </div>
                    <button class="btn btn-green" onclick="addUser()">➕ Create User</button>
                </div>
            </div>
        </div>
    </div>

    <!-- TAB 1: CLASSES -->
    <div id="tab-1" class="tab-content">
        <div class="section">
            <div class="section-title">📚 Classes / Streams</div>
            <div class="input-group"><label>Add Stream (e.g. 9A, 10B)</label><input type="text" id="classNameInput" placeholder="Enter stream name"></div>
            <button class="btn" onclick="addClass()">Add Stream</button>
            <div id="classList" style="margin-top:16px;"></div>
        </div>
    </div>

    <!-- TAB 2: SUBJECTS -->
    <div id="tab-2" class="tab-content">
        <div class="section">
            <div class="section-title">📖 Subjects &amp; Maximum Marks — All Grades at Once</div>
            <div class="info-box">
                Set maximum marks per subject for <strong>each grade in one table</strong>.<br>
                • Different grades can have different max marks for the same subject.<br>
                • Leave a cell blank to inherit the <strong>Default</strong> value.<br>
                • ENGP1+ENGP2 and KISP1+KISP2 are summed — set the <em>combined</em> max here.
            </div>

            <!-- Add / manage grade columns -->
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
                <div class="input-group" style="margin:0;flex:0 0 auto;">
                    <label style="font-size:12px;">Add Grade Column</label>
                    <input type="text" id="newGradeColInput" placeholder="e.g. 7, 8, 9" style="width:100px;padding:8px;border:1px solid #d1d5db;border-radius:4px;font-size:13px;">
                </div>
                <button class="btn" style="margin-top:18px;" onclick="addGradeColumn()">➕ Add Grade</button>
                <span style="margin-top:18px;font-size:12px;color:#6b7280;">Grades detected from streams: <span id="detectedGradesSpan" style="font-weight:700;color:#1e40af;"></span></span>
                <button class="btn btn-green" style="margin-top:18px;" onclick="autoDetectGrades()">🔍 Auto-detect from Streams</button>
            </div>

            <!-- Add subject row -->
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
                <div class="input-group" style="margin:0;flex:1;min-width:160px;">
                    <label style="font-size:12px;">Add Subject</label>
                    <input type="text" id="subjectNameInput" placeholder="e.g. Mathematics">
                </div>
                <button class="btn" style="margin-top:18px;" onclick="addSubjectRow()">➕ Add Subject</button>
            </div>

            <!-- All-grades-at-once grid -->
            <div id="subjectMaxGrid" style="margin-top:4px;"></div>

            <!-- Save all button -->
            <div style="margin-top:14px;">
                <button class="btn btn-green" onclick="saveAllSubjectMaxMarks()">💾 Save All Max Marks</button>
                <span id="saveAllSubjStatus" style="font-size:12px;color:#15803d;margin-left:10px;"></span>
            </div>
        </div>
    </div>

    <!-- TAB 3: PERFORMANCE LEVELS -->
    <div id="tab-3" class="tab-content">
        <div class="section">
            <div class="section-title">⚙️ Performance Levels</div>
            <p style="font-size:13px;color:#6b7280;margin-bottom:14px;">Grades based on percentage. Raw marks are auto-converted to % using subject max marks.</p>
            <div class="grid-2">
                <div class="input-group"><label>Grade Label (e.g. A, EE)</label><input type="text" id="gradeInput" placeholder="e.g. A"></div>
                <div class="input-group"><label>Min %</label><input type="number" id="minScoreInput" placeholder="0"></div>
                <div class="input-group"><label>Max %</label><input type="number" id="maxScoreInput" placeholder="100"></div>
            </div>
            <button class="btn" onclick="addGrade()">Add Grade</button>
            <div id="performanceList" style="margin-top:16px;"></div>
        </div>
    </div>

    <!-- TAB 4: ENTER MARKS -->
    <div id="tab-4" class="tab-content">
        <div class="section">
            <div class="section-title">📤 Enter Exam Marks Manually</div>
            <!-- Teacher assignment info (shown for teacher role only) -->
            <div id="teacherAssignmentInfo" style="display:none;background:#f0fdf4;border:1px solid #86efac;border-radius:4px;padding:10px 14px;margin-bottom:12px;font-size:13px;color:#14532d;"></div>
            <div class="grid-2">
                <div class="input-group">
                    <label>Select Active Exam</label>
                    <select id="examNameInput">
                        <option value="">-- Select Exam --</option>
                    </select>
                    <div id="examNameInfo" style="font-size:12px;color:#6b7280;margin-top:4px;"></div>
                </div>
                <div class="input-group"><label>Stream</label><select id="examClassSelect"></select></div>
                <div class="input-group"><label>Subject</label><select id="examSubjectSelect"></select></div>
            </div>
            <button class="btn" onclick="loadExamLearners()">Load Learners</button>
            <div id="marksEntryForm" style="margin-top:16px;"></div>
            <div id="saveMarksBtn" style="display:none;margin-top:12px;">
                <button class="btn btn-green" onclick="saveExamMarks()">💾 Save All Marks</button>
            </div>
        </div>
    </div>

    <!-- TAB 5: ANALYSIS -->
    <div id="tab-5" class="tab-content">
        <div class="section">
            <div class="section-title">📊 Exam Analysis</div>
            <div class="input-group" style="max-width:340px;"><label>Select Exam</label>
                <select id="analysisExam" onchange="generateAnalysis()"><option value="">-- Select Exam --</option></select>
            </div>
            <div class="mode-tabs">
                <button class="mode-tab active" onclick="setMode('overview',this)">📋 Overview</button>
                <button class="mode-tab" onclick="setMode('subject',this)">📖 By Subject</button>
                <button class="mode-tab" onclick="setMode('learner',this)">👤 Learner Rankings</button>
                <button class="mode-tab" onclick="setMode('gradecompare',this)">🏫 Grade Comparison</button>
            </div>
            <div id="analysisResults"></div>
        </div>
    </div>

    <!-- TAB 6: PAST RESULTS -->
    <div id="tab-6" class="tab-content">
        <div class="section">
            <div class="section-title">📈 Past Results</div>
            <button class="btn" onclick="loadPastResults()">🔄 Refresh</button>
            <div id="pastResultsList" style="margin-top:16px;"></div>
        </div>
    </div>

    <!-- TAB 7: DOWNLOADS -->
    <div id="tab-7" class="tab-content">
        <div class="section">
            <div class="section-title">📥 Download Reports</div>
            <div class="input-group" style="max-width:340px;"><label>Select Exam</label>
                <select id="downloadExamSelect" onchange="refreshDownloadStreams()"><option value="">-- Select Exam --</option></select>
            </div>
            <button class="btn btn-green" onclick="downloadClassAnalysis()">📥 Stream Analysis (Excel)</button>
            <button class="btn btn-green" onclick="downloadLevelDistribution()">📥 Level Distribution (Excel)</button>
            <button class="btn btn-green" onclick="downloadFullReport()">📥 Full Report — Ranked (Excel)</button>
            <button class="btn btn-green" onclick="downloadUnrankedReport()">📥 Full Report — Unranked (Excel)</button>
            <button class="btn btn-green" onclick="downloadGradeComparison()">📥 Grade Comparison (Excel)</button>
            <button class="btn btn-green" onclick="downloadGradeRanking()">📥 Grade Ranking — All Streams (Excel)</button>
            <button class="btn btn-green" onclick="downloadTopPerformers()">🥇 Top 3 Performers per Grade (Excel)</button>

            <hr style="margin:18px 0;border:none;border-top:2px solid #e2e8f0;">
            <div class="section-title" style="font-size:14px;">🏆 Most Improved Learners</div>
            <p style="font-size:12px;color:#6b7280;margin-bottom:10px;">Top 3 most improved per Grade (all streams combined), based on Total % deviation between two exams.</p>
            <div class="grid-2" style="max-width:560px;">
                <div class="input-group">
                    <label>Earlier Exam (baseline)</label>
                    <select id="improvedExam1"><option value="">-- Select Exam --</option></select>
                </div>
                <div class="input-group">
                    <label>Later Exam (comparison)</label>
                    <select id="improvedExam2"><option value="">-- Select Exam --</option></select>
                </div>
            </div>
            <button class="btn btn-green" onclick="downloadMostImproved()">🏆 Download Most Improved (Excel)</button>

            <hr style="margin:18px 0;border:none;border-top:2px solid #e2e8f0;">

            <div class="section-title" style="font-size:14px;">🧾 Individual Learner Report Cards</div>
            <p style="font-size:12px;color:#6b7280;margin-bottom:12px;">Select up to 3 exams. Score % and performance level appear per exam. Deviation column shows change between the last two exams selected.</p>
            <div class="grid-2" style="max-width:780px;">
                <div class="input-group">
                    <label>Exam 1 (oldest / first)</label>
                    <select id="rcExam1" onchange="refreshRCDropdowns()"><option value="">-- None --</option></select>
                </div>
                <div class="input-group">
                    <label>Exam 2</label>
                    <select id="rcExam2" onchange="refreshRCDropdowns()"><option value="">-- None --</option></select>
                </div>
                <div class="input-group">
                    <label>Exam 3 (most recent)</label>
                    <select id="rcExam3" onchange="refreshRCDropdowns()"><option value="">-- None --</option></select>
                </div>
                <div class="input-group">
                    <label>Filter by Stream (optional)</label>
                    <select id="reportCardStream" onchange="refreshReportLearners()">
                        <option value="">-- All Streams --</option>
                    </select>
                </div>
                <div class="input-group">
                    <label>Select Learner (or all)</label>
                    <select id="reportCardLearner">
                        <option value="">-- All Learners --</option>
                    </select>
                </div>
            </div>
            <button class="btn btn-green" onclick="downloadReportCards()">🧾 Download Report Card(s) (PDF-ready HTML)</button>
            <br><br>
            <button class="btn btn-red" onclick="clearAllData()">🗑️ Clear All Data</button>
        </div>
    </div>

    <!-- TAB 8: IMPORT LEARNERS -->
    <div id="tab-8" class="tab-content">
        <div class="section">
            <div class="section-title">👥 Import Learners from Excel</div>
            <div class="info-box">
                <strong>Excel Format:</strong><br>
                • Column A: Serial No (ignored) &nbsp;|&nbsp; Column B: Learner Name &nbsp;|&nbsp; Column C: Gender &nbsp;|&nbsp; Column D: Stream<br>
                • <strong>Rows 1–5</strong>: headers / info (skipped). Data starts from <strong>Row 6</strong>.<br>
                • Multiple streams in one file are grouped automatically.<br>
                • Workbooks with multiple sheets: pick the sheet after uploading.
            </div>
            <!-- STEP 1: file -->
            <div class="input-group"><label>Step 1 — Select Excel File (.xlsx, .xls, .csv)</label>
                <input type="file" id="learnersFile" accept=".xlsx,.xls,.csv" onchange="loadLearnersWorkbook()">
            </div>
            <!-- STEP 2: sheet picker (hidden until file loaded) -->
            <div id="learnersSheetPicker" style="display:none;">
                <div class="input-group"><label>Step 2 — Choose Sheet</label>
                    <select id="learnersSheetSelect" onchange="previewLearners()"></select>
                </div>
            </div>
            <button class="btn btn-green" onclick="importLearners()">📥 Import &amp; Save Learners</button>
            <div id="learnersPreview" style="margin-top:16px;"></div>
        </div>
    </div>

    <!-- TAB 9: IMPORT MARKS -->
    <div id="tab-9" class="tab-content">
        <div class="section">
            <div class="section-title">📊 Import Exam Marks from Excel</div>
            <div class="info-box">
                <strong>Excel Format:</strong><br>
                • Column A: Serial No (ignored) &nbsp;|&nbsp; Column B: Learner Name &nbsp;|&nbsp; Column C: Gender &nbsp;|&nbsp; Column D: Stream<br>
                • Column E onwards: Subject marks — <strong>subject names must be in Row 5</strong>, data starts from Row 6.<br>
                • ENGP1 + ENGP2 are summed → stored as <strong>English Total</strong>. KISP1 + KISP2 → <strong>Kiswahili Total</strong>.<br>
                • Blank columns between subjects are treated as separators and ignored.<br>
                • Workbooks with multiple sheets: pick the sheet after uploading.
            </div>
            <div class="input-group" style="max-width:400px;">
                <label>Select Active Exam</label>
                <select id="importExamName" onchange="onImportExamChange()">
                    <option value="">-- Select Exam --</option>
                </select>
                <div id="importExamInfo" style="font-size:12px;color:#6b7280;margin-top:4px;"></div>
            </div>
            <!-- STEP 1: file -->
            <div class="input-group"><label>Step 1 — Select Excel File (.xlsx, .xls, .csv)</label>
                <input type="file" id="marksFile" accept=".xlsx,.xls,.csv" onchange="loadMarksWorkbook()">
            </div>
            <!-- STEP 2: sheet picker -->
            <div id="marksSheetPicker" style="display:none;">
                <div class="input-group"><label>Step 2 — Choose Sheet</label>
                    <select id="marksSheetSelect" onchange="previewMarks()"></select>
                </div>
            </div>
            <button class="btn btn-green" onclick="importMarks()">📥 Import &amp; Save Marks</button>
            <div id="marksPreview" style="margin-top:16px;"></div>
        </div>
    </div>

    <!-- TAB 13: EXAMS -->
    <div id="tab-13" class="tab-content">
        <div class="section">
            <div class="section-title">📝 Exam Management</div>
            <p style="font-size:13px;color:#6b7280;margin-bottom:16px;">Create exams here. Active exams appear in the Import Marks dropdown. Close an exam when marks entry is complete.</p>

            <!-- Create exam form -->
            <div class="block" style="margin-bottom:18px;">
                <div class="block-hdr">➕ Create New Exam</div>
                <div style="padding:16px;">
                    <div class="grid-2">
                        <div class="input-group">
                            <label>Exam Name</label>
                            <input type="text" id="newExamName" placeholder="e.g. Term 1 2025, Mid Term, Final Exam">
                        </div>
                        <div class="input-group">
                            <label>Exam Type</label>
                            <select id="newExamType">
                                <option value="Term Exam">Term Exam</option>
                                <option value="Mid Term">Mid Term</option>
                                <option value="End Term">End Term</option>
                                <option value="Mock Exam">Mock Exam</option>
                                <option value="CAT">CAT</option>
                                <option value="Other">Other</option>
                            </select>
                        </div>
                        <div class="input-group">
                            <label>Academic Year</label>
                            <input type="text" id="newExamYear" placeholder="e.g. 2025">
                        </div>
                        <div class="input-group">
                            <label>Term / Semester</label>
                            <select id="newExamTerm">
                                <option value="Term 1">Term 1</option>
                                <option value="Term 2">Term 2</option>
                                <option value="Term 3">Term 3</option>
                                <option value="Semester 1">Semester 1</option>
                                <option value="Semester 2">Semester 2</option>
                            </select>
                        </div>
                    </div>
                    <button class="btn btn-green" onclick="createExamRecord()">➕ Create Exam</button>
                    <span id="createExamStatus" style="font-size:12px;color:#15803d;margin-left:8px;"></span>
                </div>
            </div>

            <!-- Exam list -->
            <div class="block">
                <div class="block-hdr">📋 All Exams</div>
                <div id="examRecordsList" style="padding:4px;"></div>
            </div>
        </div>
    </div>

    <!-- TAB 10: LEARNERS STORE -->
    <div id="tab-10" class="tab-content">
        <div class="section">
            <div class="section-title">🗂️ Learners Store</div>
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:14px;">
                <div class="input-group" style="margin:0;flex:1;min-width:180px;">
                    <select id="learnersStoreStream" onchange="renderLearnersStore()" style="padding:9px 10px;border:1px solid #d1d5db;border-radius:4px;font-size:14px;width:100%;">
                        <option value="">-- All Streams --</option>
                    </select>
                </div>
                <button class="btn btn-red" onclick="deleteAllLearnersInView()" style="margin:0;">🗑️ Delete All Shown</button>
            </div>
            <div id="learnersStoreContent"></div>
        </div>
    </div>

    <!-- TAB 11: MARKS STORE -->
    <div id="tab-11" class="tab-content">
        <div class="section">
            <div class="section-title">📑 Marks Store</div>
            <!-- Filters row -->
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px;">
                <div class="input-group" style="margin:0;flex:1;min-width:180px;">
                    <select id="marksStoreExam" onchange="renderMarksStore()" style="padding:9px 10px;border:1px solid #d1d5db;border-radius:4px;font-size:14px;width:100%;">
                        <option value="">-- Select Exam --</option>
                    </select>
                </div>
                <div class="input-group" style="margin:0;flex:1;min-width:180px;">
                    <select id="marksStoreStream" onchange="renderMarksStore()" style="padding:9px 10px;border:1px solid #d1d5db;border-radius:4px;font-size:14px;width:100%;">
                        <option value="">-- All Streams --</option>
                    </select>
                </div>
                <button class="btn btn-red" onclick="deleteAllMarksInView()" style="margin:0;">🗑️ Delete All Shown</button>
            </div>
            <!-- Sub-tab bar -->
            <div style="display:flex;gap:6px;margin-bottom:14px;border-bottom:2px solid #e2e8f0;padding-bottom:0;">
                <button id="msTabView" class="mode-tab active" onclick="switchMarksStoreTab('view')" style="border-radius:4px 4px 0 0;margin-bottom:-2px;">👁 View Marks</button>
                <button id="msTabEdit" class="mode-tab" onclick="switchMarksStoreTab('edit')" style="border-radius:4px 4px 0 0;margin-bottom:-2px;">✏️ Edit Marks</button>
            </div>
            <!-- View pane -->
            <div id="msViewPane">
                <div id="marksStoreContent"></div>
            </div>
            <!-- Edit pane -->
            <div id="msEditPane" style="display:none;">
                <div class="info-box" style="margin-bottom:12px;">
                    Select exam and stream above, then edit marks directly in the cells below.<br>
                    Click <strong>💾 Save Edits</strong> to save, or <strong>📊 Save &amp; Analyse</strong> to save then jump to Analysis.
                </div>
                <div id="marksEditContent"></div>
            </div>
        </div>
    </div>

    <!-- TAB 12: TEACHERS -->
    <div id="tab-12" class="tab-content">
        <div class="section">
            <div class="section-title">👨‍🏫 Teachers — Class Teachers & Principal</div>
            <p style="font-size:13px;color:#6b7280;margin-bottom:16px;">Assign class teachers to streams and set the principal. Their names and uploaded signatures will appear on individual report cards.</p>

            <!-- Principal -->
            <div class="block" style="margin-bottom:18px;">
                <div class="block-hdr">🎓 Principal</div>
                <div style="padding:16px;" class="grid-2">
                    <div>
                        <div class="input-group"><label>Principal's Name</label>
                            <input type="text" id="principalName" placeholder="e.g. Mr. John Kamau">
                        </div>
                        <div class="input-group"><label>Title / Designation</label>
                            <input type="text" id="principalTitle" placeholder="e.g. B.Ed, M.Ed" value="Principal">
                        </div>
                        <button class="btn btn-green" onclick="savePrincipal()">💾 Save Principal</button>
                        <span id="principalStatus" style="font-size:12px;color:#15803d;margin-left:8px;"></span>
                    </div>
                    <div>
                        <div class="input-group"><label>Principal's Signature (image)</label>
                            <input type="file" id="principalSigFile" accept="image/*" onchange="uploadSignature('principal')">
                        </div>
                        <div id="principalSigPreview" style="margin-top:8px;"></div>
                    </div>
                </div>
            </div>

            <!-- Class Teachers -->
            <div class="block">
                <div class="block-hdr">📚 Class Teachers — Assign to Streams</div>
                <div style="padding:16px;">
                    <div class="grid-2">
                        <div class="input-group"><label>Stream</label>
                            <select id="teacherStream">
                                <option value="">-- Select Stream --</option>
                            </select>
                        </div>
                        <div class="input-group"><label>Teacher's Name</label>
                            <input type="text" id="teacherName" placeholder="e.g. Ms. Jane Wanjiku">
                        </div>
                        <div class="input-group"><label>Link to System Username (optional)</label>
                            <select id="teacherUsername">
                                <option value="">-- Select username --</option>
                            </select>
                            <div style="font-size:11px;color:#6b7280;margin-top:3px;">Links this teacher's login to their stream for filtered mark entry.</div>
                        </div>
                        <div class="input-group"><label>Assigned Subjects (comma-separated, blank = all)</label>
                            <input type="text" id="teacherSubjects" placeholder="e.g. Mathematics, Science">
                            <div style="font-size:11px;color:#6b7280;margin-top:3px;">Leave blank to allow all subjects.</div>
                        </div>
                        <div class="input-group"><label>Signature (image, optional)</label>
                            <input type="file" id="teacherSigFile" accept="image/*" onchange="uploadSignature('teacher')">
                        </div>
                        <div id="teacherSigPreview" style="margin-top:8px;"></div>
                    </div>
                    <button class="btn btn-green" onclick="saveClassTeacher()">💾 Assign Teacher to Stream</button>
                    <div id="teachersList" style="margin-top:18px;"></div>
                </div>
            </div>
        </div>
    </div>

    <!-- TAB 14: IMPORT ALL-IN-ONE -->
    <div id="tab-14" class="tab-content">
        <div class="section">
            <div class="section-title">⚡ Import All-in-One — Learners &amp; Marks Together</div>
            <div class="info-box">
                <strong>How it works:</strong><br>
                • Upload one Excel workbook where <strong>each sheet = one stream's data</strong>.<br>
                • The system detects all sheets and lets you assign each sheet to a stream.<br>
                • <strong>Row 5</strong>: column headers — Col A: Serial No &nbsp;|&nbsp; Col B: Name &nbsp;|&nbsp; Col C: Gender &nbsp;|&nbsp; Col D onwards: Subject marks (ENGP1, ENGP2, KISP1, KISP2, etc.)<br>
                • <strong>Row 6+</strong>: learner data.<br>
                • Learners are imported and marks saved in one step.
            </div>

            <!-- Step 1: Select exam -->
            <div class="input-group" style="max-width:420px;">
                <label>Step 1 — Select Active Exam for Marks</label>
                <select id="aioExamName" onchange="aioOnExamChange()">
                    <option value="">-- Select Exam --</option>
                </select>
                <div id="aioExamInfo" style="font-size:12px;color:#6b7280;margin-top:4px;"></div>
            </div>

            <!-- Step 2: Upload file -->
            <div class="input-group">
                <label>Step 2 — Upload Excel Workbook (.xlsx, .xls)</label>
                <input type="file" id="aioFile" accept=".xlsx,.xls,.csv" onchange="aioLoadWorkbook()">
            </div>

            <!-- Step 3: Sheet → Stream mapping (shown after file loaded) -->
            <div id="aioSheetMapper" style="display:none;">
                <div style="font-weight:bold;font-size:13px;color:#1e3a8a;margin-bottom:10px;">
                    Step 3 — Assign each sheet to a stream
                </div>
                <div id="aioMappingRows"></div>
                <div id="aioPreviewArea" style="margin-top:16px;"></div>
                <div style="margin-top:14px;">
                    <button class="btn btn-green" onclick="aioImportAll()">⚡ Import All — Learners &amp; Marks</button>
                    <span id="aioImportStatus" style="font-size:13px;font-weight:600;margin-left:10px;"></span>
                </div>
            </div>
        </div>
    </div>

</div><!-- /container -->

<script>
// ═══════════════ DATA ═══════════════
var schoolSetup = {name:''};
var classes  = [];   // ["9A","9B",...]
var subjects = [];   // [{name, maxMarks}]
var grades   = [];   // [{grade, minScore, maxScore}] sorted desc
var learners = {};   // {stream: [{name,gender}]}
var exams    = {};   // {examName: ["9A_Maths","9B_Maths",...]}
var examMarks= {};   // {"examName_stream_subject": {learnerName: rawMark}}
// examRecords: created exam metadata [{id,name,type,year,term,status,createdAt}]
var examRecords = [];

// gradeSubjectConfig: per-grade max marks override
// Structure: { "default":{subjects:{name:max}, eng:max, kis:max}, "9":{...}, "10":{...} }
// "default" applies to any grade not explicitly configured
var gradeSubjectConfig = { default:{ subjects:{}, eng:100, kis:100 } };

var analysisMode = 'overview';

// ═══════════════ UTILS ═══════════════
function showMessage(text, type) {
    var m=document.getElementById('message');
    m.textContent=text; m.className='message show '+type;
    clearTimeout(m._t); m._t=setTimeout(function(){ m.classList.remove('show'); },4500);
}
function getRole()     { try{ return JSON.parse(sessionStorage.getItem('csSession')||'{}').role||'teacher'; }catch(e){ return 'teacher'; } }
function getUsername() { try{ return JSON.parse(sessionStorage.getItem('csSession')||'{}').username||''; }catch(e){ return ''; } }
function isAdmin()     { return getRole()==='admin'; }

// Tabs allowed for each role
// teacher: Enter Marks (4), Analysis (5), Past Results (6)
var TEACHER_TABS = [4, 5, 6];

function switchTab(i, btn) {
    // Role check
    if (!isAdmin() && !TEACHER_TABS.includes(i)) {
        showMessage('⛔ Access denied — your role cannot access this section.', 'error');
        return;
    }
    document.querySelectorAll('.tab-content').forEach(function(t){ t.classList.remove('active'); });
    document.querySelectorAll('.tab').forEach(function(t){ t.classList.remove('active'); });
    document.getElementById('tab-'+i).classList.add('active');
    if (btn) btn.classList.add('active');
    else {
        document.querySelectorAll('.tab').forEach(function(b){
            if (b.getAttribute('onclick')==='switchTab('+i+',this)') b.classList.add('active');
        });
    }
    if(i===0) renderUserManager();
    if(i===2){ refreshSubjectList(); openSubjectsTab(); }
    if(i===3) refreshPerformanceList();
    if(i===4) refreshExamSelects();
    if(i===5){ refreshAllDropdowns(); generateAnalysis(); }
    if(i===6) { refreshAllDropdowns(); loadPastResults(); }
    if(i===7) { refreshAllDropdowns(); refreshDownloadExams(); }
    if(i===10) openLearnersStore();
    if(i===11) { refreshAllDropdowns(); openMarksStore(); }
    if(i===12) openTeachersTab();
    if(i===13) openExamsTab();
    if(i===14) aioOpenTab();
}

// Show/hide tabs and buttons based on role
function applyRoleUI() {
    var admin = isAdmin();

    // Tab bar — show only allowed tabs for teachers
    document.querySelectorAll('.tab').forEach(function(btn){
        var onclick = btn.getAttribute('onclick')||'';
        var m = onclick.match(/switchTab\((\d+)/);
        if (!m) return;
        var tabIdx = parseInt(m[1]);
        if (!admin && !TEACHER_TABS.includes(tabIdx)) {
            btn.style.display = 'none';
        } else {
            btn.style.display = '';
        }
    });

    // Analysis tab (5) — hide download/edit controls for teachers
    // Past Results (6) — already read-only (no delete shown for teachers below)

    // Marks Store edit tab — disabled for teachers (handled in renderMarksEdit)
    var msTabEdit = document.getElementById('msTabEdit');
    if (msTabEdit) msTabEdit.style.display = admin ? '' : 'none';

    // Enter Marks — for teachers: filter to assigned stream only (done in refreshExamSelects)
    // Downloads tab buttons — hide for teachers (they can't download from Tab 7 since tab is hidden)
}

// Teacher: get their assigned streams by matching login username to teachers.streams[stream].username
function getTeacherStreams() {
    var username = getUsername().toLowerCase();
    var assigned = [];
    Object.entries(teachers.streams||{}).forEach(function(e){
        var stream=e[0], t=e[1];
        if (t.username && t.username.toLowerCase()===username) assigned.push(stream);
    });
    return assigned.length ? assigned : null; // null = all streams allowed
}

// Teacher: get assigned subjects for a specific stream
function getTeacherSubjects(stream) {
    if (isAdmin()) return null; // null = all subjects
    var t = (teachers.streams||{})[stream];
    if (!t || !t.subjects || !t.subjects.length) return null;
    return t.subjects; // specific list
}
// Refresh all exam/stream dropdowns across tabs
function refreshAllDropdowns() {
    refreshAnalysisExams();
    refreshDownloadExams();
    refreshImportExamDropdown();
    refreshExamSelects(); // Tab 4 Enter Marks exam dropdown
    var examSel=document.getElementById('marksStoreExam');
    if(examSel){
        var cur=examSel.value;
        examSel.innerHTML='<option value="">-- Select Exam --</option>';
        Object.keys(exams).forEach(function(e){
            var o=document.createElement('option'); o.value=e; o.textContent=e;
            if(e===cur) o.selected=true; examSel.appendChild(o);
        });
    }
}
// Navigate programmatically to analysis tab
function goToAnalysis() {
    switchTab(5, null);
}
function setMode(mode,btn) {
    analysisMode=mode;
    document.querySelectorAll('.mode-tab').forEach(function(b){ b.classList.remove('active'); });
    btn.classList.add('active'); generateAnalysis();
}

// ═══════════════ PAPER CONFIG (English & Kiswahili) ═══════════════
// paperConfig = { eng:{p1Max,p2Max}, kis:{p1Max,p2Max} }
var paperConfig = {eng:{p1Max:50,p2Max:50}, kis:{p1Max:50,p2Max:50}};

function togglePaperConfig(lang) {
    var el=document.getElementById(lang+'PaperConfig');
    el.style.display = el.style.display==='none' ? 'block' : 'none';
}
function updatePaperSummary(lang) {
    var p1=parseFloat(document.getElementById(lang+'P1Max').value)||0;
    var el=document.getElementById(lang+'PaperSummary');
    el.textContent = p1 ? 'Max: '+p1 : '';
}
function savePaperConfig(lang) {
    var p1=parseFloat(document.getElementById(lang+'P1Max').value);
    if(isNaN(p1)||p1<=0){ showMessage('❌ Enter valid maximum marks','error'); return; }
    paperConfig[lang]={p1Max:p1, p2Max:0};
    saveAllData();
    updatePaperSummary(lang);
    var label=lang==='eng'?'English':'Kiswahili';
    document.getElementById(lang+'SaveStatus').textContent='✓ Saved (max='+p1+')';
    setTimeout(function(){ document.getElementById(lang+'SaveStatus').textContent=''; },3000);
    showMessage('✓ '+label+' max marks saved ('+p1+')','success');
}
function loadPaperConfigUI() {
    ['eng','kis'].forEach(function(lang){
        var cfg=paperConfig[lang]||{p1Max:100,p2Max:0};
        document.getElementById(lang+'P1Max').value=cfg.p1Max||'';
        updatePaperSummary(lang);
    });
}

// Get max marks for a subject, optionally for a specific grade number
// Lookup: grade-specific config → default config → subjects array → fallback 100
function getMaxMarks(subjectName, gradeNum) {
    var n = String(subjectName).trim().toUpperCase();
    var isEng = (n==='ENGLISH TOTAL'||n==='ENGLISH'||n==='ENG');
    var isKis = (n==='KISWAHILI TOTAL'||n==='KISWAHILI'||n==='KIS');

    // Try grade-specific config first, then default
    var configs = [];
    if (gradeNum && gradeSubjectConfig[String(gradeNum)]) configs.push(gradeSubjectConfig[String(gradeNum)]);
    if (gradeSubjectConfig['default']) configs.push(gradeSubjectConfig['default']);

    for (var ci=0; ci<configs.length; ci++) {
        var cfg = configs[ci];
        if (isEng && cfg.eng) return cfg.eng;
        if (isKis && cfg.kis) return cfg.kis;
        // Check subjects map
        if (cfg.subjects) {
            var subjMax = cfg.subjects[subjectName] || cfg.subjects[subjectName.toUpperCase()];
            if (!subjMax) {
                // case-insensitive
                Object.keys(cfg.subjects).forEach(function(k){
                    if (k.toUpperCase()===n) subjMax=cfg.subjects[k];
                });
            }
            if (subjMax) return subjMax;
        }
    }

    // Fallback: global subjects array
    var subj = subjects.find(function(s){ return s.name===subjectName; });
    if (subj) return subj.maxMarks || 100;
    subj = subjects.find(function(s){ return s.name.toUpperCase()===n; });
    if (subj) return subj.maxMarks || 100;
    // fuzzy
    subj = subjects.find(function(s){
        var sn=s.name.toUpperCase();
        return sn.indexOf(n)>=0 || n.indexOf(sn)>=0;
    });
    if (subj) return subj.maxMarks || 100;

    return 100; // fallback
}

// ═══════════════ GRADING ═══════════════
// toPercent: pass stream so we can extract grade for grade-aware max marks
function toPercent(raw, subjectName, stream) {
    var gradeNum = stream ? gradeFromStream(stream) : null;
    return (parseFloat(raw) / getMaxMarks(subjectName, gradeNum)) * 100;
}
function getGrade(pct) {
    if(!grades.length) return null;
    for(var i=0;i<grades.length;i++){
        if(pct>=grades[i].minScore&&pct<=grades[i].maxScore) return grades[i].grade;
    }
    return '—';
}
function gradeBadge(g) {
    if(!g||g==='—') return '<span class="badge badge-x">—</span>';
    var idx=grades.findIndex(function(gr){ return gr.grade===g; });
    var cls=idx===0?'badge-0':idx===1?'badge-1':idx===2?'badge-2':idx>=3?'badge-3':'badge-x';
    return '<span class="badge '+cls+'">'+g+'</span>';
}
function calcStats(arr) {
    if(!arr||!arr.length) return null;
    var sum=arr.reduce(function(a,b){ return a+b; },0);
    return {
        count:arr.length, avg:sum/arr.length,
        highest:Math.max.apply(null,arr), lowest:Math.min.apply(null,arr),
        pass:arr.filter(function(p){ return p>=50; }).length,
        fail:arr.filter(function(p){ return p<50; }).length
    };
}

// Extract grade number from stream name e.g. "9A" -> "9", "10B" -> "10"
function gradeFromStream(stream) {
    var m=String(stream).match(/^(\d+)/);
    return m ? m[1] : stream;
}

// Parse "9A_Mathematics" -> {stream:"9A", subject:"Mathematics"}
function parseCombo(combo) {
    var idx=combo.indexOf('_');
    if(idx<0) return {stream:combo,subject:''};
    return {stream:combo.substring(0,idx), subject:combo.substring(idx+1)};
}

// ═══════════════ TAB 0: SCHOOL ═══════════════
function saveSchoolData() {
    schoolSetup.name    = document.getElementById('schoolName').value.trim();
    schoolSetup.motto   = document.getElementById('schoolMotto').value.trim();
    schoolSetup.address = document.getElementById('schoolAddress').value.trim();
    saveAllData(); updateSchoolDisplay();
    showMessage('✓ School information saved','success');
}
function updateSchoolDisplay() {
    document.getElementById('schoolDisplay').textContent = 'School: '+(schoolSetup.name||'Not Set');
    document.getElementById('schoolName').value    = schoolSetup.name   ||'';
    document.getElementById('schoolMotto').value   = schoolSetup.motto  ||'';
    document.getElementById('schoolAddress').value = schoolSetup.address||'';
    if(schoolSetup.logo){
        document.getElementById('logoPreview').innerHTML =
            '<img src="'+schoolSetup.logo+'" style="max-height:80px;max-width:180px;border:1px solid #e2e8f0;border-radius:4px;padding:4px;">';
        document.getElementById('logoStatus').textContent = '✓ Logo saved';
    }
}
function uploadLogo() {
    var file=document.getElementById('logoFile').files[0]; if(!file) return;
    if(file.size>500000){ showMessage('❌ Logo image too large — keep under 500 KB','error'); return; }
    var reader=new FileReader();
    reader.onload=function(e){
        schoolSetup.logo=e.target.result; saveAllData();
        document.getElementById('logoPreview').innerHTML=
            '<img src="'+e.target.result+'" style="max-height:80px;max-width:180px;border:1px solid #e2e8f0;border-radius:4px;padding:4px;">';
        document.getElementById('logoStatus').textContent='✓ Logo saved';
        showMessage('✓ School logo uploaded and saved','success');
    };
    reader.readAsDataURL(file);
}

// ═══════════════ TAB 12: TEACHERS ═══════════════
// teachers = { principal:{name,title,sig}, streams:{"9A":{name,sig}, ...} }
var teachers = { principal:{name:'',title:'Principal',sig:''}, streams:{} };
var _pendingTeacherSig = ''; // temp hold uploaded sig before saving

function openTeachersTab() {
    var sel=document.getElementById('teacherStream');
    var cur=sel.value;
    sel.innerHTML='<option value="">-- Select Stream --</option>';
    classes.slice().sort().forEach(function(c){
        var o=document.createElement('option'); o.value=c; o.textContent=c;
        if(c===cur) o.selected=true; sel.appendChild(o);
    });
    // Populate username dropdown with teacher-role users
    var uSel=document.getElementById('teacherUsername');
    if(uSel){
        var curU=uSel.value;
        uSel.innerHTML='<option value="">-- Select username --</option>';
        getUsers().filter(function(u){ return u.role==='teacher'; }).forEach(function(u){
            var o=document.createElement('option'); o.value=u.username; o.textContent=u.username;
            if(u.username===curU) o.selected=true; uSel.appendChild(o);
        });
    }
    document.getElementById('principalName').value  = teachers.principal.name||'';
    document.getElementById('principalTitle').value = teachers.principal.title||'Principal';
    renderPrincipalSig();
    renderTeachersList();
    _pendingTeacherSig='';
}

function saveClassTeacher() {
    var stream  =document.getElementById('teacherStream').value;
    var name    =document.getElementById('teacherName').value.trim();
    var username=document.getElementById('teacherUsername').value;
    var subjsRaw=document.getElementById('teacherSubjects').value.trim();
    if(!stream){ showMessage('❌ Select a stream','error'); return; }
    if(!name)  { showMessage('❌ Enter teacher name','error'); return; }
    var subjList=subjsRaw ? subjsRaw.split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [];
    if(!teachers.streams[stream]) teachers.streams[stream]={name:'',sig:'',username:'',subjects:[]};
    teachers.streams[stream].name    =name;
    teachers.streams[stream].username=username;
    teachers.streams[stream].subjects=subjList;
    if(_pendingTeacherSig) teachers.streams[stream].sig=_pendingTeacherSig;
    saveAllData(); _pendingTeacherSig='';
    document.getElementById('teacherName').value='';
    document.getElementById('teacherSubjects').value='';
    document.getElementById('teacherSigFile').value='';
    document.getElementById('teacherSigPreview').innerHTML='';
    renderTeachersList();
    showMessage('✓ Teacher assigned to '+stream,'success');
}

function uploadSignature(who) {
    var fileId = who==='principal' ? 'principalSigFile' : 'teacherSigFile';
    var file=document.getElementById(fileId).files[0]; if(!file) return;
    if(file.size>300000){ showMessage('❌ Signature image too large — keep under 300 KB','error'); return; }
    var reader=new FileReader();
    reader.onload=function(e){
        if(who==='principal'){
            teachers.principal.sig=e.target.result; saveAllData();
            renderPrincipalSig();
            showMessage('✓ Principal signature saved','success');
        } else {
            _pendingTeacherSig=e.target.result;
            document.getElementById('teacherSigPreview').innerHTML=
                '<img src="'+e.target.result+'" style="max-height:50px;border:1px solid #e2e8f0;border-radius:3px;">'
                +'<span style="font-size:11px;color:#6b7280;margin-left:6px;">Signature ready — click Assign Teacher</span>';
        }
    };
    reader.readAsDataURL(file);
}

function renderPrincipalSig() {
    var sig=teachers.principal.sig;
    document.getElementById('principalSigPreview').innerHTML = sig
        ? '<img src="'+sig+'" style="max-height:60px;max-width:160px;border:1px solid #e2e8f0;border-radius:3px;padding:3px;">'
          +'<div style="font-size:11px;color:#15803d;margin-top:3px;">✓ Signature saved</div>'
          +'<button onclick="clearPrincipalSig()" style="font-size:11px;background:#fee2e2;color:#dc2626;border:1px solid #fca5a5;border-radius:3px;padding:2px 8px;cursor:pointer;margin-top:4px;">Remove</button>'
        : '<span style="font-size:12px;color:#9ca3af;">No signature uploaded</span>';
}
function clearPrincipalSig(){
    teachers.principal.sig=''; saveAllData(); renderPrincipalSig();
    showMessage('✓ Principal signature removed','success');
}

function savePrincipal() {
    teachers.principal.name  = document.getElementById('principalName').value.trim();
    teachers.principal.title = document.getElementById('principalTitle').value.trim()||'Principal';
    saveAllData();
    document.getElementById('principalStatus').textContent='✓ Saved';
    setTimeout(function(){ document.getElementById('principalStatus').textContent=''; },2500);
    showMessage('✓ Principal information saved','success');
}

function renderTeachersList() {
    var streams=Object.keys(teachers.streams);
    if(!streams.length){
        document.getElementById('teachersList').innerHTML='<p style="font-size:13px;color:#6b7280;">No class teachers assigned yet.</p>';
        return;
    }
    var html='<div class="tbl-wrap"><table><tr><th>Stream</th><th>Name</th><th>Username</th><th>Subjects</th><th>Signature</th><th></th></tr>';
    streams.sort().forEach(function(st){
        var t=teachers.streams[st];
        var subjs=(t.subjects&&t.subjects.length)?t.subjects.join(', '):'All subjects';
        var uname=t.username||'—';
        html+='<tr>'
            +'<td><span class="stream-tag">'+st+'</span></td>'
            +'<td><strong>'+t.name+'</strong></td>'
            +'<td><span style="font-size:12px;background:#eff6ff;color:#1e3a8a;padding:2px 7px;border-radius:4px;">'+uname+'</span></td>'
            +'<td style="font-size:12px;color:#374151;">'+subjs+'</td>'
            +'<td>'+(t.sig?'<img src="'+t.sig+'" style="max-height:36px;max-width:100px;border:1px solid #e2e8f0;border-radius:2px;">':'<span style="color:#9ca3af;font-size:12px;">None</span>')+'</td>'
            +'<td><button class="item-delete" onclick="removeTeacher(\''+st+'\')">Remove</button></td></tr>';
    });
    html+='</table></div>';
    document.getElementById('teachersList').innerHTML=html;
}

function removeTeacher(stream) {
    if(!confirm('Remove teacher from stream '+stream+'?')) return;
    delete teachers.streams[stream]; saveAllData(); renderTeachersList();
    showMessage('✓ Teacher removed from '+stream,'success');
}

// ═══════════════ TAB 1 ═══════════════
function addClass() {
    var n=document.getElementById('classNameInput').value.trim();
    if(!n){ showMessage('❌ Enter a stream name','error'); return; }
    if(classes.includes(n)){ showMessage('❌ Stream already exists','error'); return; }
    classes.push(n); if(!learners[n]) learners[n]=[];
    saveAllData(); document.getElementById('classNameInput').value='';
    refreshClassList(); showMessage('✓ Stream added: '+n,'success');
}
function refreshClassList() {
    if(!classes.length){ document.getElementById('classList').innerHTML='<p style="color:#6b7280;font-size:13px;">No streams yet.</p>'; return; }
    var html='';
    // group by grade
    var byGrade={};
    classes.forEach(function(c){ var g=gradeFromStream(c); if(!byGrade[g]) byGrade[g]=[]; byGrade[g].push(c); });
    Object.keys(byGrade).sort().forEach(function(g){
        html+='<p style="font-weight:bold;color:#1e3a8a;margin:10px 0 5px;">Grade '+g+'</p>';
        byGrade[g].forEach(function(cls,i){
            var idx=classes.indexOf(cls);
            var cnt=(learners[cls]||[]).length;
            html+='<div class="item"><span><span class="stream-tag">'+cls+'</span> <span style="color:#6b7280;font-size:12px;">'+cnt+' learner'+(cnt!==1?'s':'')+'</span></span>'
                +'<button class="item-delete" onclick="deleteClass('+idx+')">Delete</button></div>';
        });
    });
    document.getElementById('classList').innerHTML=html;
}
function deleteClass(idx) {
    if(!confirm('Delete stream '+classes[idx]+'?')) return;
    delete learners[classes[idx]]; classes.splice(idx,1);
    saveAllData(); refreshClassList(); showMessage('✓ Stream deleted','success');
}

// ═══════════════ TAB 2: SUBJECTS — ALL-GRADES-AT-ONCE GRID ═══════════════
var _gradeColumns = []; // grade keys shown as columns: ['default','7','8','9',...]

function openSubjectsTab() {
    if (!gradeSubjectConfig['default']) gradeSubjectConfig['default'] = {subjects:{},eng:100,kis:100};
    _gradeColumns = Object.keys(gradeSubjectConfig).sort(function(a,b){
        if(a==='default') return -1; if(b==='default') return 1; return parseInt(a)-parseInt(b);
    });
    if (!_gradeColumns.includes('default')) _gradeColumns.unshift('default');
    updateDetectedGradesSpan();
    renderSubjectMaxGrid();
}

function updateDetectedGradesSpan() {
    var detected=[...new Set(classes.map(gradeFromStream))].filter(Boolean).sort(function(a,b){ return parseInt(a)-parseInt(b); });
    var el=document.getElementById('detectedGradesSpan');
    if(el) el.textContent=detected.length?detected.join(', '):'(none yet — add streams first)';
}

function autoDetectGrades() {
    var detected=[...new Set(classes.map(gradeFromStream))].filter(Boolean);
    var added=0;
    detected.forEach(function(g){
        if(!_gradeColumns.includes(g)){
            if(!gradeSubjectConfig[g]){
                var def=gradeSubjectConfig['default']||{};
                gradeSubjectConfig[g]={subjects:Object.assign({},def.subjects||{}),eng:def.eng||100,kis:def.kis||100};
            }
            _gradeColumns.push(g); added++;
        }
    });
    _gradeColumns.sort(function(a,b){ if(a==='default') return -1; if(b==='default') return 1; return parseInt(a)-parseInt(b); });
    saveAllData(); updateDetectedGradesSpan(); renderSubjectMaxGrid();
    showMessage(added?'✓ Added '+added+' grade column(s)':'All detected grades already have columns', added?'success':'info');
}

function addGradeColumn() {
    var val=document.getElementById('newGradeColInput').value.trim();
    if(!val){ showMessage('❌ Enter a grade number','error'); return; }
    if(_gradeColumns.includes(val)){ showMessage('Grade '+val+' already exists','info'); return; }
    if(!gradeSubjectConfig[val]){
        var def=gradeSubjectConfig['default']||{};
        gradeSubjectConfig[val]={subjects:Object.assign({},def.subjects||{}),eng:def.eng||100,kis:def.kis||100};
    }
    _gradeColumns.push(val);
    _gradeColumns.sort(function(a,b){ if(a==='default') return -1; if(b==='default') return 1; return parseInt(a)-parseInt(b); });
    document.getElementById('newGradeColInput').value='';
    saveAllData(); renderSubjectMaxGrid();
    showMessage('✓ Grade '+val+' column added','success');
}

function removeGradeColumn(g) {
    if(g==='default'){ showMessage('❌ Cannot remove Default column','error'); return; }
    if(!confirm('Remove Grade '+g+' config? Max marks will fall back to Default for all subjects in Grade '+g+'.')) return;
    delete gradeSubjectConfig[g];
    _gradeColumns=_gradeColumns.filter(function(x){ return x!==g; });
    saveAllData(); renderSubjectMaxGrid();
    showMessage('✓ Grade '+g+' removed — falling back to Default values','success');
}

function addSubjectRow() {
    var name=document.getElementById('subjectNameInput').value.trim();
    if(!name){ showMessage('❌ Enter a subject name','error'); return; }
    var up=name.toUpperCase();
    if(up==='ENGLISH'||up==='ENGLISH TOTAL'||up==='KISWAHILI'||up==='KISWAHILI TOTAL'){
        showMessage('❌ English and Kiswahili are always in the grid — set their max marks there','error'); return;
    }
    if(subjects.find(function(s){ return s.name===name; })){ showMessage('Subject "'+name+'" already exists','info'); document.getElementById('subjectNameInput').value=''; renderSubjectMaxGrid(); return; }
    // Add to all grade configs with default 100
    _gradeColumns.forEach(function(g){
        if(!gradeSubjectConfig[g]) gradeSubjectConfig[g]={subjects:{},eng:100,kis:100};
        if(!gradeSubjectConfig[g].subjects[name]) gradeSubjectConfig[g].subjects[name]=100;
    });
    subjects.push({name:name,maxMarks:100});
    document.getElementById('subjectNameInput').value='';
    saveAllData(); renderSubjectMaxGrid();
    showMessage('✓ Subject "'+name+'" added','success');
}

function deleteSubjectRow(name) {
    if(!confirm('Remove subject "'+name+'" from all grade configs?')) return;
    _gradeColumns.forEach(function(g){
        if(gradeSubjectConfig[g]&&gradeSubjectConfig[g].subjects) delete gradeSubjectConfig[g].subjects[name];
    });
    subjects=subjects.filter(function(s){ return s.name!==name; });
    saveAllData(); renderSubjectMaxGrid();
    showMessage('✓ Subject "'+name+'" removed','success');
}

function encodeGridId(name){ return name.replace(/[^a-zA-Z0-9]/g,'_'); }

function renderSubjectMaxGrid() {
    var container=document.getElementById('subjectMaxGrid'); if(!container) return;

    // Collect all other subject names
    var allSubjNames=new Set();
    _gradeColumns.forEach(function(g){ Object.keys((gradeSubjectConfig[g]||{}).subjects||{}).forEach(function(n){ allSubjNames.add(n); }); });
    subjects.forEach(function(s){ if(s.name!=='English Total'&&s.name!=='Kiswahili Total') allSubjNames.add(s.name); });
    ['English','Kiswahili','English Total','Kiswahili Total'].forEach(function(n){ allSubjNames.delete(n); });

    // Sort by SUBJECT_ORDER
    var otherNames=[...allSubjNames].sort(function(a,b){ return subjectSortKey(a)-subjectSortKey(b); });

    var thS='padding:10px 12px;text-align:center;font-size:12px;white-space:nowrap;';
    var tdS='padding:8px 8px;text-align:center;border-bottom:1px solid #e5e7eb;';
    var inpS='width:78px;padding:5px;border:1px solid #d1d5db;border-radius:3px;text-align:center;font-size:13px;';

    var html='<div class="tbl-wrap"><table style="border-collapse:collapse;width:100%;font-size:13px;">';
    // Header
    html+='<thead><tr>'
        +'<th style="'+thS+'text-align:left;background:#1e3a8a;color:white;min-width:200px;">Subject</th>';
    _gradeColumns.forEach(function(g){
        var isD=g==='default';
        html+='<th style="'+thS+'background:'+(isD?'#374151':'#1e40af')+';color:white;min-width:110px;">'
            +(isD?'⭐ Default':'Grade '+g)
            +(isD?'':'<br><button onclick="removeGradeColumn(\''+g+'\')" style="font-size:9px;background:rgba(255,255,255,0.2);border:none;color:white;border-radius:3px;padding:1px 5px;cursor:pointer;margin-top:3px;">✕ Remove</button>')
            +'</th>';
    });
    html+='</tr></thead><tbody>';

    // Fixed rows
    [{prop:'eng',label:'🇬🇧 English (ENGP1+ENGP2)',bg:'#f0fdf4'},{prop:'kis',label:'🇰🇪 Kiswahili (KISP1+KISP2)',bg:'#dcfce7'}].forEach(function(fr){
        html+='<tr style="background:'+fr.bg+';">'
            +'<td style="'+tdS+'text-align:left;font-weight:700;color:#15803d;">'+fr.label+'</td>';
        _gradeColumns.forEach(function(g){
            var cfg=gradeSubjectConfig[g]||{};
            var val=cfg[fr.prop]||'';
            var defVal=(gradeSubjectConfig['default']||{})[fr.prop]||100;
            html+='<td style="'+tdS+'background:'+(g==='default'?fr.bg:'white')+';">'
                +'<input type="number" id="grid_'+fr.prop+'_'+g+'" value="'+val+'" min="1" '
                +'placeholder="'+(g==='default'?'100':defVal)+'" style="'+inpS+'">'
                +(g!=='default'&&!val?'<div style="font-size:9px;color:#9ca3af;margin-top:2px;">Default: '+defVal+'</div>':'')
                +'</td>';
        });
        html+='</tr>';
    });

    // Other subject rows
    otherNames.forEach(function(name,ri){
        html+='<tr style="background:'+(ri%2===0?'#ffffff':'#f8fafc')+';">'
            +'<td style="'+tdS+'text-align:left;font-weight:600;">'+name
            +' <button onclick="deleteSubjectRow(\''+name.replace(/'/g,"\\'")+'\')" style="font-size:10px;background:#fee2e2;color:#dc2626;border:1px solid #fca5a5;border-radius:3px;padding:1px 6px;cursor:pointer;margin-left:4px;">✕</button>'
            +'</td>';
        _gradeColumns.forEach(function(g){
            var cfg=gradeSubjectConfig[g]||{};
            var val=(cfg.subjects&&cfg.subjects[name])?cfg.subjects[name]:'';
            var defVal=((gradeSubjectConfig['default']||{}).subjects||{})[name]||100;
            html+='<td style="'+tdS+'background:'+(g==='default'?'#fafafa':'white')+';">'
                +'<input type="number" id="grid_subj_'+encodeGridId(name)+'_'+g+'" value="'+val+'" min="1" '
                +'placeholder="'+(g==='default'?'100':defVal)+'" style="'+inpS+'">'
                +(g!=='default'&&!val?'<div style="font-size:9px;color:#9ca3af;margin-top:2px;">Default: '+defVal+'</div>':'')
                +'</td>';
        });
        html+='</tr>';
    });
    if(!otherNames.length){
        html+='<tr><td colspan="'+(1+_gradeColumns.length)+'" style="padding:12px;color:#9ca3af;font-size:12px;text-align:center;">No other subjects yet — add one above.</td></tr>';
    }
    html+='</tbody></table></div>';
    container.innerHTML=html;
}

function saveAllSubjectMaxMarks() {
    // Save English + Kiswahili
    [{prop:'eng'},{prop:'kis'}].forEach(function(fr){
        _gradeColumns.forEach(function(g){
            var inp=document.getElementById('grid_'+fr.prop+'_'+g); if(!inp) return;
            var val=parseFloat(inp.value);
            if(isNaN(val)||val<=0) return;
            if(!gradeSubjectConfig[g]) gradeSubjectConfig[g]={subjects:{},eng:100,kis:100};
            gradeSubjectConfig[g][fr.prop]=val;
            // Keep legacy paperConfig in sync
            if(fr.prop==='eng') paperConfig.eng.p1Max=val;
            if(fr.prop==='kis') paperConfig.kis.p1Max=val;
        });
    });
    // Save other subjects
    var allSubjNames=new Set();
    _gradeColumns.forEach(function(g){ Object.keys((gradeSubjectConfig[g]||{}).subjects||{}).forEach(function(n){ allSubjNames.add(n); }); });
    subjects.forEach(function(s){ if(s.name!=='English Total'&&s.name!=='Kiswahili Total') allSubjNames.add(s.name); });
    ['English','Kiswahili','English Total','Kiswahili Total'].forEach(function(n){ allSubjNames.delete(n); });

    allSubjNames.forEach(function(name){
        _gradeColumns.forEach(function(g){
            var inp=document.getElementById('grid_subj_'+encodeGridId(name)+'_'+g); if(!inp) return;
            var val=parseFloat(inp.value);
            if(isNaN(val)||val<=0) return;
            if(!gradeSubjectConfig[g]) gradeSubjectConfig[g]={subjects:{},eng:100,kis:100};
            if(!gradeSubjectConfig[g].subjects) gradeSubjectConfig[g].subjects={};
            gradeSubjectConfig[g].subjects[name]=val;
            // Sync with global subjects array
            var ex=subjects.find(function(s){ return s.name===name; });
            if(!ex) subjects.push({name:name,maxMarks:val});
            else if(g==='default') ex.maxMarks=val;
        });
    });

    saveAllData();
    var el=document.getElementById('saveAllSubjStatus');
    if(el){ el.textContent='✓ All saved!'; setTimeout(function(){ el.textContent=''; },3000); }
    renderSubjectMaxGrid();
    showMessage('✓ All subject max marks saved for all grades','success');
}

// Legacy stubs (kept so old code references don't break)
function refreshSubjectList(){ renderSubjectMaxGrid(); }
function addSubject(){ addSubjectRow(); }
function deleteSubject(idx){ subjects.splice(idx,1); saveAllData(); renderSubjectMaxGrid(); }
function renderAllGradesConfigSummary(){} // replaced by grid
function deleteGradeConfig(g){ removeGradeColumn(g); }
function refreshSubjectGradeSelect(){}
function addGradeToSubjects(){ var n=window.prompt('Grade number:',''); if(n) { document.getElementById('newGradeColInput').value=n.trim(); addGradeColumn(); } }
function switchSubjectGrade(){}
function saveSubjectMaxForGrade(){ saveAllSubjectMaxMarks(); }

// ═══════════════ TAB 3 ═══════════════
function addGrade() {
    var grade=document.getElementById('gradeInput').value.trim();
    var mn=parseFloat(document.getElementById('minScoreInput').value);
    var mx=parseFloat(document.getElementById('maxScoreInput').value);
    if(!grade||isNaN(mn)||isNaN(mx)){ showMessage('❌ Fill all fields','error'); return; }
    if(mn>=mx){ showMessage('❌ Min must be less than Max','error'); return; }
    grades.push({grade:grade,minScore:mn,maxScore:mx});
    grades.sort(function(a,b){ return b.minScore-a.minScore; });
    saveAllData();
    document.getElementById('gradeInput').value='';
    document.getElementById('minScoreInput').value='';
    document.getElementById('maxScoreInput').value='';
    refreshPerformanceList(); showMessage('✓ Grade added: '+grade,'success');
}
function refreshPerformanceList() {
    if(!grades.length){ document.getElementById('performanceList').innerHTML='<p style="color:#6b7280;font-size:13px;">No grades defined.</p>'; return; }
    var html='<div class="tbl-wrap"><table><tr><th>Grade</th><th>Min %</th><th>Max %</th><th></th></tr>';
    grades.forEach(function(g,idx){
        html+='<tr><td>'+gradeBadge(g.grade)+' <strong>'+g.grade+'</strong></td><td>'+g.minScore+'%</td><td>'+g.maxScore+'%</td><td><button class="item-delete" onclick="deleteGrade('+idx+')">Delete</button></td></tr>';
    });
    document.getElementById('performanceList').innerHTML=html+'</table></div>';
}
function deleteGrade(idx){ grades.splice(idx,1); saveAllData(); refreshPerformanceList(); }

// ═══════════════ TAB 4 ═══════════════
function refreshExamSelects() {
    // Exam dropdown — active exams only
    var examSel = document.getElementById('examNameInput');
    if (examSel && examSel.tagName === 'SELECT') {
        var curExam = examSel.value;
        examSel.innerHTML = '<option value="">-- Select Active Exam --</option>';
        var activeExams = examRecords.filter(function(r){ return r.status==='active'; });
        if (!activeExams.length) {
            var o = document.createElement('option');
            o.disabled = true; o.textContent = '(No active exams — create one in Exams tab)';
            examSel.appendChild(o);
        } else {
            activeExams.forEach(function(r){
                var o = document.createElement('option');
                o.value = r.name;
                o.textContent = r.name + ' · ' + r.type + ' · ' + r.term + ' ' + r.year;
                if (r.name === curExam) o.selected = true;
                examSel.appendChild(o);
            });
        }
        // Show exam info below dropdown
        examSel.onchange = function() {
            var info = document.getElementById('examNameInfo');
            if (!info) return;
            var val = examSel.value;
            var r = examRecords.find(function(r){ return r.name===val; });
            if (r) {
                var cnt = 0;
                (exams[r.name]||[]).forEach(function(c){ cnt += Object.keys(examMarks[r.name+'_'+c]||{}).length; });
                info.textContent = r.type + ' · ' + r.term + ' ' + r.year + (cnt ? ' · ' + cnt + ' marks already entered' : ' · No marks yet');
                info.style.color = cnt ? '#b45309' : '#15803d';
            } else { info.textContent = ''; }
        };
    }
    // Stream and Subject dropdowns — filtered by teacher assignment if teacher role
    var role=getRole(), username=getUsername();
    var assignedStreams=[], assignedSubjects=[];
    if(role==='teacher'){
        Object.entries(teachers.streams||{}).forEach(function(e){
            var st=e[0], t=e[1];
            if(t.username===username||t.name===username){
                assignedStreams.push(st);
                if(t.subjects&&t.subjects.length) t.subjects.forEach(function(s){ if(!assignedSubjects.includes(s)) assignedSubjects.push(s); });
            }
        });
    }
    var cs=document.getElementById('examClassSelect'), ss=document.getElementById('examSubjectSelect');
    cs.innerHTML='<option value="">-- Select Stream --</option>';
    ss.innerHTML='<option value="">-- Select Subject --</option>';
    var streamList=(role==='teacher'&&assignedStreams.length)?assignedStreams:classes;
    streamList.forEach(function(c){ var o=document.createElement('option');o.value=c;o.textContent=c;cs.appendChild(o); });
    var subjList=(role==='teacher'&&assignedSubjects.length)?subjects.filter(function(s){ return assignedSubjects.includes(s.name); }):subjects;
    subjList.forEach(function(s){ var o=document.createElement('option');o.value=s.name;o.textContent=s.name+' (max '+s.maxMarks+')';ss.appendChild(o); });
    // Teacher assignment info panel
    var taInfo=document.getElementById('teacherAssignmentInfo');
    if(taInfo){
        if(role==='teacher'){
            if(assignedStreams.length){
                taInfo.innerHTML='📋 Your assignment: Stream(s) <strong>'+assignedStreams.join(', ')+'</strong>'+(assignedSubjects.length?'&nbsp;|&nbsp;Subjects: <strong>'+assignedSubjects.join(', ')+'</strong>':'&nbsp;|&nbsp;All subjects');
                taInfo.style.display='block';
            } else {
                taInfo.innerHTML='⚠ You have not been assigned to any stream yet. Contact an admin to configure your assignment in the Teachers tab.';
                taInfo.style.display='block'; taInfo.style.background='#fef2f2'; taInfo.style.borderColor='#fca5a5';
            }
        } else { taInfo.style.display='none'; }
    }
}
function loadExamLearners() {
    var examName=document.getElementById('examNameInput').value;
    var stream=document.getElementById('examClassSelect').value;
    var subject=document.getElementById('examSubjectSelect').value;
    if(!examName){ showMessage('❌ Select an active exam — create one in the Exams tab first','error'); return; }
    if(!stream||!subject){ showMessage('❌ Select a stream and subject','error'); return; }
    var subj=subjects.find(function(s){ return s.name===subject; });
    var maxM=subj?subj.maxMarks:100;
    var cls_learners=learners[stream]||[];
    if(!cls_learners.length){ showMessage('⚠️ No learners in this stream. Import learners first.','info'); return; }
    var key=examName+'_'+stream+'_'+subject;
    var existing=examMarks[key]||{};
    var html='<p style="font-size:13px;color:#6b7280;margin-bottom:10px;">Enter raw marks out of <strong>'+maxM+'</strong>.</p>';
    html+='<div class="tbl-wrap"><table><tr><th>#</th><th>Learner</th><th>Gender</th><th>Raw Marks (out of '+maxM+')</th><th>%</th><th>Grade</th></tr>';
    cls_learners.forEach(function(l,idx){
        var prev=existing[l.name]!==undefined?existing[l.name]:'';
        var pct=prev!==''?((prev/maxM)*100).toFixed(0):'';
        var gr=pct!==''?getGrade(parseFloat(pct)):'';
        html+='<tr><td>'+(idx+1)+'</td><td>'+l.name+'</td><td>'+l.gender+'</td>'
            +'<td><input type="number" id="mark_'+idx+'" value="'+prev+'" placeholder="0" min="0" max="'+maxM+'" '
            +'oninput="livePreview('+idx+','+maxM+')" style="width:80px;padding:5px;border:1px solid #d1d5db;border-radius:3px;"></td>'
            +'<td id="pct_'+idx+'" style="color:#1e40af;font-weight:bold;">'+(pct?pct+'%':'')+'</td>'
            +'<td id="grd_'+idx+'">'+(gr?gradeBadge(gr):'')+'</td></tr>';
    });
    html+='</table></div>';
    var form=document.getElementById('marksEntryForm');
    form.innerHTML=html; form.dataset.stream=stream; form.dataset.subject=subject;
    document.getElementById('saveMarksBtn').style.display='block';
    showMessage('✓ Loaded '+cls_learners.length+' learners','success');
}
function livePreview(idx,maxM) {
    var raw=parseFloat(document.getElementById('mark_'+idx).value);
    var pc=document.getElementById('pct_'+idx), gc=document.getElementById('grd_'+idx);
    if(!isNaN(raw)){ var pct=(raw/maxM)*100; pc.textContent=pct.toFixed(0)+'%'; gc.innerHTML=gradeBadge(getGrade(pct)); }
    else { pc.textContent=''; gc.innerHTML=''; }
}
function saveExamMarks() {
    var examName=document.getElementById('examNameInput').value;
    var form=document.getElementById('marksEntryForm');
    var stream=form.dataset.stream, subject=form.dataset.subject;
    if(!examName||!stream||!subject){ showMessage('❌ Load learners first','error'); return; }
    var cls_learners=learners[stream]||[];
    var key=examName+'_'+stream+'_'+subject, marks={}, count=0;
    cls_learners.forEach(function(l,idx){
        var v=document.getElementById('mark_'+idx).value;
        if(v!==''){ marks[l.name]=parseFloat(v); count++; }
    });
    if(!count){ showMessage('❌ Enter at least one mark','error'); return; }
    examMarks[key]=marks;
    if(!exams[examName]) exams[examName]=[];
    var combo=stream+'_'+subject;
    if(!exams[examName].includes(combo)) exams[examName].push(combo);
    saveAllData(); showMessage('✓ Saved marks for '+count+' learners','success');
}

// ═══════════════ TAB 5: ANALYSIS ═══════════════
function refreshAnalysisExams() {
    var sel=document.getElementById('analysisExam'), cur=sel.value;
    sel.innerHTML='<option value="">-- Select Exam --</option>';
    Object.keys(exams).forEach(function(e){
        var o=document.createElement('option'); o.value=e; o.textContent=e;
        if(e===cur) o.selected=true; sel.appendChild(o);
    });
}
function generateAnalysis() {
    var examName=document.getElementById('analysisExam').value;
    var container=document.getElementById('analysisResults');
    if(!examName){ container.innerHTML=''; return; }
    var combos=exams[examName]||[];
    if(!combos.length){ container.innerHTML='<p style="color:#6b7280;">No data for this exam.</p>'; return; }
    if(analysisMode==='overview')      renderOverview(examName,combos,container);
    else if(analysisMode==='subject')  renderBySubject(examName,combos,container);
    else if(analysisMode==='learner')  renderByLearner(examName,combos,container);
    else                               renderGradeComparison(examName,combos,container);
}

// ── Helpers ──
// Fixed subject display order: Maths, English, Kiswahili, Pre-Tech, Agri, CAS, CRE, Science, Soc Studies
var SUBJECT_ORDER = [
    'Mathematics','Maths','Math','MATH',
    'English Total','English',
    'Kiswahili Total','Kiswahili',
    'Pre-Technical Studies','Pre-Technical','Pre-Tech','PRE-TECH','Pre Tech','Pretechnical',
    'Agriculture','AGRI','Agri',
    'CAS',
    'CRE',
    'Science','INTSC','Int Science','Integrated Science',
    'Social Studies','SOST','Soc Studies','SocSt'
];

function subjectSortKey(s) {
    var up = s.toUpperCase();
    for (var i = 0; i < SUBJECT_ORDER.length; i++) {
        if (SUBJECT_ORDER[i].toUpperCase() === up) return i;
    }
    // Fuzzy fallback
    if (up.indexOf('MATH') >= 0) return 0;
    if (up === 'ENGLISH TOTAL' || up.indexOf('ENG') >= 0) return 10;
    if (up === 'KISWAHILI TOTAL' || up.indexOf('KIS') >= 0 || up.indexOf('KISW') >= 0) return 20;
    if (up.indexOf('PRE') >= 0 || up.indexOf('TECH') >= 0) return 30;
    if (up.indexOf('AGRI') >= 0) return 40;
    if (up.indexOf('CAS') >= 0) return 50;
    if (up.indexOf('CRE') >= 0) return 60;
    if (up.indexOf('SCI') >= 0 || up.indexOf('INTSC') >= 0) return 70;
    if (up.indexOf('SOC') >= 0 || up.indexOf('SOST') >= 0) return 80;
    return 999;
}

// Paper subjects: English Total and Kiswahili Total — graded, highlighted
function isPaperSubject(s){ return s==='English Total'||s==='Kiswahili Total'; }

// Raw mark → % for a subject
function comboPcts(examName,combo) {
    var p=parseCombo(combo), key=examName+'_'+combo;
    return Object.values(examMarks[key]||{}).map(function(r){ return toPercent(r,p.subject,p.stream); });
}

// Subject display label
function subjLabel(s){
    if(s==='English Total') return 'English';
    if(s==='Kiswahili Total') return 'Kiswahili';
    return s;
}

// Filter combos: exclude raw paper columns
function cleanCombos(combos){
    return combos.filter(function(c){
        var s=parseCombo(c).subject.toUpperCase();
        return s!=='ENGP1'&&s!=='ENGP2'&&s!=='KISP1'&&s!=='KISP2';
    });
}

// Get ordered list of unique subjects from combos using fixed order
function orderedSubjects(combos){
    var subs=[...new Set(combos.map(function(c){
        return c.indexOf('_')>=0 ? parseCombo(c).subject : c;
    }))];
    return subs.sort(function(a,b){ return subjectSortKey(a)-subjectSortKey(b); });
}

// ── OVERVIEW ──
function renderOverview(examName,combos,container) {
    combos=cleanCombos(combos);
    var allPcts=[], subjMap={}, streamMap={};
    combos.forEach(function(combo){
        var p=parseCombo(combo), pcts=comboPcts(examName,combo);
        pcts.forEach(function(pct){
            allPcts.push(pct);
            if(!subjMap[p.subject]) subjMap[p.subject]=[];
            subjMap[p.subject].push(pct);
            if(!streamMap[p.stream]) streamMap[p.stream]=[];
            streamMap[p.stream].push(pct);
        });
    });
    var s=calcStats(allPcts);
    if(!s){ container.innerHTML='<p>No marks data.</p>'; return; }
    var html='<div class="stat-cards">'
        +'<div class="stat-card"><div class="val">'+s.count+'</div><div class="lbl">Total Entries</div></div>'
        +'<div class="stat-card"><div class="val">'+s.avg.toFixed(2)+'%</div><div class="lbl">Overall Avg</div></div>'
        +'<div class="stat-card"><div class="val">'+s.highest.toFixed(0)+'%</div><div class="lbl">Highest</div></div>'
        +'<div class="stat-card"><div class="val">'+s.lowest.toFixed(0)+'%</div><div class="lbl">Lowest</div></div>'
        +'<div class="stat-card"><div class="val">'+s.pass+'</div><div class="lbl">Above 50%</div></div>'
        +'<div class="stat-card"><div class="val">'+s.fail+'</div><div class="lbl">Below 50%</div></div>'
        +'</div>';
    if(grades.length){
        html+='<h4 style="margin:12px 0 8px;color:#1e3a8a;">Overall Grade Distribution</h4>';
        html+='<div class="tbl-wrap"><table><tr><th>Grade</th><th>Count</th><th>% of Total</th></tr>';
        grades.forEach(function(g){
            var cnt=allPcts.filter(function(p){ return p>=g.minScore&&p<=g.maxScore; }).length;
            html+='<tr><td>'+gradeBadge(g.grade)+' '+g.grade+'</td><td>'+cnt+'</td><td>'+((cnt/s.count)*100).toFixed(0)+'%</td></tr>';
        });
        html+='</table></div>';
    }
    // Subject averages — % and grade badge in same cell for all subjects
    html+='<h4 style="margin:12px 0 8px;color:#1e3a8a;">Subject Averages</h4>';
    html+='<div class="tbl-wrap"><table><tr><th>Subject</th><th>Max Marks</th><th>Entries</th><th>Avg % &amp; Level</th><th>Highest</th><th>Lowest</th><th>Pass</th><th>Fail</th></tr>';
    orderedSubjects(combos).forEach(function(subj){
        var pcts=subjMap[subj]||[]; var ss=calcStats(pcts); if(!ss) return;
        var isP=isPaperSubject(subj);
        var maxM=getMaxMarks(subj);
        html+='<tr'+(isP?' style="background:#f0fdf4;"':'')+'>'
            +'<td><strong>'+subjLabel(subj)+'</strong>'+(isP?' <span style="font-size:10px;background:#dcfce7;color:#15803d;padding:1px 5px;border-radius:8px;">P1+P2</span>':'')+'</td>'
            +'<td>'+maxM+'</td>'
            +'<td>'+ss.count+'</td>'
            +'<td><strong>'+parseFloat(ss.avg.toFixed(2))+'%</strong> &nbsp;'+gradeBadge(getGrade(ss.avg))+'</td>'
            +'<td>'+ss.highest.toFixed(0)+'%</td><td>'+ss.lowest.toFixed(0)+'%</td>'
            +'<td style="color:#15803d;">'+ss.pass+'</td><td style="color:#dc2626;">'+ss.fail+'</td></tr>';
    });
    html+='</table></div>';
    html+='<h4 style="margin:12px 0 8px;color:#1e3a8a;">Stream Averages</h4>';
    html+='<div class="tbl-wrap"><table><tr><th>Stream</th><th>Grade No.</th><th>Entries</th><th>Avg %</th><th>Highest</th><th>Lowest</th><th>Grade</th></tr>';
    Object.entries(streamMap).sort(function(a,b){ return gradeFromStream(a[0]).localeCompare(gradeFromStream(b[0]))||a[0].localeCompare(b[0]); }).forEach(function(e){
        var ss=calcStats(e[1]);
        html+='<tr><td><span class="stream-tag">'+e[0]+'</span></td><td>'+gradeFromStream(e[0])+'</td><td>'+ss.count+'</td>'
            +'<td>'+parseFloat(ss.avg.toFixed(2))+'%</td><td>'+ss.highest.toFixed(0)+'%</td><td>'+ss.lowest.toFixed(0)+'%</td>'
            +'<td>'+gradeBadge(getGrade(ss.avg))+'</td></tr>';
    });
    html+='</table></div>';
    container.innerHTML=html;
}

// ── BY SUBJECT ──
function renderBySubject(examName,combos,container) {
    combos=cleanCombos(combos);
    var bySubject={};
    combos.forEach(function(combo){
        var p=parseCombo(combo);
        if(!bySubject[p.subject]) bySubject[p.subject]=[];
        bySubject[p.subject].push(p.stream);
    });
    var html='';
    orderedSubjects(combos).forEach(function(subject){
        var streams=bySubject[subject]||[], subjPcts=[];
        streams.forEach(function(st){ comboPcts(examName,st+'_'+subject).forEach(function(p){ subjPcts.push(p); }); });
        var ss=calcStats(subjPcts); if(!ss) return;
        var isP=isPaperSubject(subject);
        var maxM=getMaxMarks(subject);
        html+='<div class="block"><div class="block-hdr" style="'+(isP?'background:#f0fdf4;':'')+'">'
            +'<span><strong>'+subjLabel(subject)+'</strong>'+(isP?' <span style="font-size:10px;background:#dcfce7;color:#15803d;padding:1px 5px;border-radius:8px;margin-left:4px;">P1+P2</span> Max: '+maxM:'  Max: '+maxM)+'</span>'
            +'<span>Overall Avg: <strong>'+parseFloat(ss.avg.toFixed(2))+'%</strong> &nbsp;'+gradeBadge(getGrade(ss.avg))+'</span></div>';
        html+='<div class="tbl-wrap"><table><tr><th>Stream</th><th>Learners</th><th>Avg % &amp; Level</th><th>Highest</th><th>Lowest</th><th>Pass</th><th>Fail</th></tr>';
        streams.sort().forEach(function(st){
            var pcts=comboPcts(examName,st+'_'+subject);
            var cs=calcStats(pcts); if(!cs) return;
            html+='<tr><td><span class="stream-tag">'+st+'</span></td><td>'+cs.count+'</td>'
                +'<td><strong>'+cs.avg.toFixed(2)+'%</strong> &nbsp;'+gradeBadge(getGrade(cs.avg))+'</td>'
                +'<td>'+cs.highest.toFixed(0)+'%</td><td>'+cs.lowest.toFixed(0)+'%</td>'
                +'<td style="color:#15803d;">'+cs.pass+'</td><td style="color:#dc2626;">'+cs.fail+'</td></tr>';
        });
        html+='</table></div>';
        if(grades.length){
            html+='<div style="padding:8px 14px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:12px;"><strong style="color:#6b7280;">GRADE DIST: </strong>';
            grades.forEach(function(g){
                var cnt=subjPcts.filter(function(p){ return p>=g.minScore&&p<=g.maxScore; }).length;
                html+=gradeBadge(g.grade)+' <span>'+cnt+'</span>&nbsp;&nbsp;';
            });
            html+='</div>';
        }
        html+='</div>';
    });
    container.innerHTML=html||'<p style="color:#6b7280;">No data.</p>';
}

// Sum of all subject percentages — e.g. 9 subjects × 100% = 900
function sumPcts(pctObj) {
    return Object.values(pctObj).filter(function(v){ return v !== undefined && !isNaN(v); })
                 .reduce(function(a,b){ return a+b; }, 0);
}

// Mean of subject percentages — sumPcts / count
function meanPct(pctObj) {
    var vals = Object.values(pctObj).filter(function(v){ return v !== undefined && !isNaN(v); });
    if (!vals.length) return 0;
    return vals.reduce(function(a,b){ return a+b; }, 0) / vals.length;
}

// ── LEARNER RANKINGS ──
function renderByLearner(examName,combos,container) {
    combos=cleanCombos(combos);
    var learnerData={};
    combos.forEach(function(combo){
        var p=parseCombo(combo), key=examName+'_'+combo;
        Object.entries(examMarks[key]||{}).forEach(function(e){
            var name=e[0], raw=e[1];
            if(!learnerData[name]) learnerData[name]={stream:p.stream,subjects:{}};
            var pct=toPercent(raw,p.subject,p.stream);
            learnerData[name].subjects[p.subject]={raw:raw,pct:pct};
        });
    });
    var allSubjects=orderedSubjects(combos);
    // Sort by mean of subject percentages descending
    var sorted=Object.entries(learnerData).sort(function(a,b){
        return meanPct(Object.fromEntries(Object.entries(a[1].subjects).map(function(e){ return [e[0],e[1].pct]; })))
             - meanPct(Object.fromEntries(Object.entries(b[1].subjects).map(function(e){ return [e[0],e[1].pct]; })));
    }).reverse();

    var html='<div class="tbl-wrap"><table><tr>'
        +'<th>Rank</th><th>Learner</th><th>Stream</th>';
    allSubjects.forEach(function(s){
        var maxM=getMaxMarks(s);
        var isP=isPaperSubject(s);
        html+='<th>'+subjLabel(s)+' /'+maxM+(isP?' <span style="font-size:9px;color:#15803d;">P1+P2</span>':'')+'</th>';
    });
    html+='<th style="background:#fef9c3;color:#713f12;">Sum</th>'
        +'<th style="background:#eff6ff;">Total %</th>'
        +'<th>Level</th></tr>';

    sorted.forEach(function(entry,idx){
        var name=entry[0], d=entry[1];
        var pctMap=Object.fromEntries(Object.entries(d.subjects).map(function(e){ return [e[0],e[1].pct]; }));
        var sum=Math.round(sumPcts(pctMap));
        var totalPct=Math.round(meanPct(pctMap));
        var totalGrade=getGrade(totalPct);
        html+='<tr>'
            +'<td style="font-weight:bold;color:#6b7280;">'+(idx+1)+'</td>'
            +'<td><strong>'+name+'</strong></td>'
            +'<td><span class="stream-tag">'+d.stream+'</span></td>';
        allSubjects.forEach(function(s){
            var isP=isPaperSubject(s);
            var sd=d.subjects[s];
            if(!sd){
                html+='<td style="color:#d1d5db;text-align:center;">—</td>';
            } else {
                var g=getGrade(sd.pct);
                html+='<td style="'+(isP?'background:#f0fdf4;':'')+'">'
                    +'<strong>'+Math.round(sd.pct)+'%</strong> '
                    +gradeBadge(g)+'</td>';
            }
        });
        html+='<td style="font-weight:bold;background:#fef9c3;color:#713f12;text-align:center;">'+sum+'</td>'
            +'<td style="font-weight:bold;background:#eff6ff;">'+totalPct+'%</td>'
            +'<td>'+gradeBadge(totalGrade)+'</td>'
            +'</tr>';
    });
    html+='</table></div>';
    container.innerHTML=html;
}

// ── GRADE COMPARISON ──
function renderGradeComparison(examName,combos,container) {
    combos=cleanCombos(combos);
    var byGrade={};
    combos.forEach(function(combo){
        var p=parseCombo(combo), g=gradeFromStream(p.stream);
        if(!byGrade[g]) byGrade[g]={streams:{},subjects:[]};
        if(!byGrade[g].streams[p.stream]) byGrade[g].streams[p.stream]={};
        if(!byGrade[g].subjects.includes(p.subject)) byGrade[g].subjects.push(p.subject);
    });
    combos.forEach(function(combo){
        var p=parseCombo(combo), g=gradeFromStream(p.stream), key=examName+'_'+combo;
        byGrade[g].streams[p.stream][p.subject]=Object.values(examMarks[key]||{}).map(function(r){ return toPercent(r,p.subject,p.stream); });
    });
    var html='';
    Object.keys(byGrade).sort().forEach(function(g){
        var gd=byGrade[g], streamNames=Object.keys(gd.streams).sort();
        var ordSubjs=orderedSubjects(gd.subjects.map(function(s){ return {subject:s}; }).map(function(o){ return '_'+o.subject; }).map(parseCombo));
        // fix: get subjects from gd.subjects directly
        var paperFirst=gd.subjects.filter(isPaperSubject);
        var otherSubjs=gd.subjects.filter(function(s){ return !isPaperSubject(s); }).sort();
        var orderedS=paperFirst.concat(otherSubjs);

        html+='<div class="block"><div class="block-hdr">🏫 Grade '+g+' — Stream Comparison'
            +' <span style="font-size:12px;">'+streamNames.map(function(s){ return '<span class="stream-tag">'+s+'</span>'; }).join(' ')+'</span></div>';
        orderedS.forEach(function(subject){
            var isP=isPaperSubject(subject);
            var maxM=getMaxMarks(subject);
            html+='<div class="block-sub"><strong>'+subjLabel(subject)+'</strong> — Max: '+maxM+(isP?' (P1+P2 combined)':'')+'</div>';
            html+='<div class="tbl-wrap"><table><tr><th>Stream</th><th>Learners</th><th>Avg % &amp; Level</th><th>Highest</th><th>Lowest</th><th>Pass</th><th>Fail</th>';
            if(grades.length) grades.forEach(function(gr){ html+='<th>'+gr.grade+'</th>'; });
            html+='</tr>';
            streamNames.forEach(function(st){
                var pcts=(gd.streams[st][subject])||[];
                var ss=calcStats(pcts);
                if(!ss){ html+='<tr><td><span class="stream-tag">'+st+'</span></td><td colspan="6" style="color:#9ca3af;">No data</td></tr>'; return; }
                html+='<tr><td><span class="stream-tag">'+st+'</span></td><td>'+ss.count+'</td>'
                    +'<td><strong>'+parseFloat(ss.avg.toFixed(2))+'%</strong> &nbsp;'+gradeBadge(getGrade(ss.avg))+'</td>'
                    +'<td>'+ss.highest.toFixed(0)+'%</td><td>'+ss.lowest.toFixed(0)+'%</td>'
                    +'<td style="color:#15803d;">'+ss.pass+'</td><td style="color:#dc2626;">'+ss.fail+'</td>';
                if(grades.length) grades.forEach(function(gr){
                    var cnt=pcts.filter(function(p){ return p>=gr.minScore&&p<=gr.maxScore; }).length;
                    html+='<td>'+cnt+'</td>';
                });
                html+='</tr>';
            });
            html+='</table></div>';
        });
        html+='<div class="block-sub">📊 Overall (all subjects combined)</div>';
        html+='<div class="tbl-wrap"><table><tr><th>Stream</th><th>Entries</th><th>Avg %</th><th>Highest</th><th>Lowest</th><th>Grade</th></tr>';
        streamNames.forEach(function(st){
            var allPcts=[];
            orderedS.forEach(function(subj){ (gd.streams[st][subj]||[]).forEach(function(p){ allPcts.push(p); }); });
            var ss=calcStats(allPcts); if(!ss) return;
            html+='<tr><td><span class="stream-tag">'+st+'</span></td><td>'+ss.count+'</td><td><strong>'+parseFloat(ss.avg.toFixed(2))+'%</strong></td>'
                +'<td>'+ss.highest.toFixed(0)+'%</td><td>'+ss.lowest.toFixed(0)+'%</td>'
                +'<td>'+gradeBadge(getGrade(ss.avg))+'</td></tr>';
        });
        html+='</table></div></div>';
    });
    container.innerHTML=html||'<p style="color:#6b7280;">No data.</p>';
}

// ═══════════════ TAB 6 ═══════════════
function loadPastResults() {
    var names=Object.keys(exams);
    var admin=isAdmin();
    if(!names.length){ document.getElementById('pastResultsList').innerHTML='<p style="color:#6b7280;font-size:13px;">No exams saved yet.</p>'; return; }
    var html='<div class="tbl-wrap"><table><tr><th>Exam Name</th><th>Streams</th><th>Combos</th><th>Total Mark Entries</th>'+(admin?'<th>Action</th>':'')+'</tr>';
    names.forEach(function(exam){
        var combos=exams[exam]||[], total=0, streams=new Set();
        combos.forEach(function(c){ total+=Object.keys(examMarks[exam+'_'+c]||{}).length; streams.add(parseCombo(c).stream); });
        html+='<tr><td><strong>'+exam+'</strong></td><td>'+streams.size+'</td><td>'+combos.length+'</td><td>'+total+'</td>'
            +(admin?'<td><button class="item-delete" onclick="deleteExam(\''+exam+'\')">Delete</button></td>':'')+'</tr>';
    });
    document.getElementById('pastResultsList').innerHTML=html+'</table></div>';
    if(!admin){
        var note=document.createElement('p');
        note.style.cssText='font-size:12px;color:#b45309;margin-top:8px;';
        note.textContent='👁 View only — contact an admin to delete exams.';
        document.getElementById('pastResultsList').appendChild(note);
    }
}
function deleteExam(examName) {
    if(!confirm('Delete exam "'+examName+'"?')) return;
    (exams[examName]||[]).forEach(function(c){ delete examMarks[examName+'_'+c]; });
    delete exams[examName]; saveAllData(); loadPastResults();
    showMessage('✓ Exam deleted','success');
}

// ═══════════════ TAB 7: DOWNLOADS ═══════════════
function refreshDownloadExams() {
    var examNames = Object.keys(exams);
    // Main download select
    var sel=document.getElementById('downloadExamSelect');
    var cur=sel.value;
    sel.innerHTML='<option value="">-- Select Exam --</option>';
    examNames.forEach(function(e){ var o=document.createElement('option');o.value=e;o.textContent=e;if(e===cur)o.selected=true;sel.appendChild(o); });
    // RC exam selects 1-3
    ['rcExam1','rcExam2','rcExam3'].forEach(function(id){
        var s=document.getElementById(id); var c=s.value;
        s.innerHTML='<option value="">-- None --</option>';
        examNames.forEach(function(e){ var o=document.createElement('option');o.value=e;o.textContent=e;if(e===c)o.selected=true;s.appendChild(o); });
    });
    // Most Improved selects
    ['improvedExam1','improvedExam2'].forEach(function(id){
        var s=document.getElementById(id); if(!s) return;
        var c=s.value;
        s.innerHTML='<option value="">-- Select Exam --</option>';
        examNames.forEach(function(e){ var o=document.createElement('option');o.value=e;o.textContent=e;if(e===c)o.selected=true;s.appendChild(o); });
    });
    refreshRCDropdowns();
}
// Called when any RC exam select changes — rebuild stream + learner dropdowns from all selected exams
function refreshRCDropdowns() {
    var selectedExams = ['rcExam1','rcExam2','rcExam3'].map(function(id){ return document.getElementById(id).value; }).filter(Boolean);
    var streamSel = document.getElementById('reportCardStream');
    var curStream = streamSel.value;
    streamSel.innerHTML = '<option value="">-- All Streams --</option>';
    var streams = new Set();
    selectedExams.forEach(function(exam){
        (exams[exam]||[]).forEach(function(c){ streams.add(parseCombo(c).stream); });
    });
    [...streams].sort().forEach(function(s){
        var o=document.createElement('option'); o.value=s; o.textContent=s;
        if(s===curStream) o.selected=true; streamSel.appendChild(o);
    });
    refreshReportLearners();
}
// Kept for backward compat (called from old onchange on downloadExamSelect)
function refreshDownloadStreams() { refreshRCDropdowns(); }

function refreshReportLearners() {
    var selectedExams = ['rcExam1','rcExam2','rcExam3'].map(function(id){ return document.getElementById(id).value; }).filter(Boolean);
    var filterStream = document.getElementById('reportCardStream').value;
    var sel = document.getElementById('reportCardLearner');
    sel.innerHTML = '<option value="">-- All Learners --</option>';
    if (!selectedExams.length) return;
    // Use first available exam to get learner list
    var names = new Set();
    selectedExams.forEach(function(exam){
        cleanCombos(exams[exam]||[]).forEach(function(combo){
            var p = parseCombo(combo);
            if (filterStream && p.stream !== filterStream) return;
            Object.keys(examMarks[exam+'_'+combo]||{}).forEach(function(n){ names.add(n+'||'+p.stream); });
        });
    });
    [...names].sort().forEach(function(entry){
        var parts=entry.split('||');
        var o=document.createElement('option'); o.value=entry;
        o.textContent=parts[0]+' ('+parts[1]+')'; sel.appendChild(o);
    });
}
function getDownloadExam() {
    var v=document.getElementById('downloadExamSelect').value;
    if(!v){ showMessage('❌ Select an exam first','error'); return null; }
    return v;
}
// Build school info header rows prepended to every sheet
// Returns array of rows: school name, motto, address, exam name, generated date
function schoolHeaderRows(examLabel) {
    var rows = [];
    rows.push([schoolSetup.name || 'SCHOOL NAME']);
    if (schoolSetup.motto)   rows.push([schoolSetup.motto]);
    if (schoolSetup.address) rows.push([schoolSetup.address]);
    rows.push(['Exam: ' + (examLabel || '')]);
    rows.push(['Generated: ' + new Date().toLocaleString()]);
    rows.push([]); // blank separator
    return rows;
}

// Export to Excel — every sheet gets school header rows prepended and column widths set
function exportXLSX(sheets, filename, examLabel) {
    var wb = XLSX.utils.book_new();
    sheets.forEach(function(s) {
        // Prepend school header rows
        var header = schoolHeaderRows(examLabel || s.examLabel || '');
        var fullData = header.concat(s.data);
        var ws = XLSX.utils.aoa_to_sheet(fullData);

        // Style: auto column widths based on content
        var maxCols = fullData.reduce(function(m, r){ return Math.max(m, r.length); }, 0);
        var colWidths = [];
        for (var c = 0; c < maxCols; c++) {
            var maxW = 10;
            fullData.forEach(function(row){
                var cell = row[c];
                if (cell !== undefined && cell !== null) {
                    var len = String(cell).length;
                    if (len > maxW) maxW = len;
                }
            });
            colWidths.push({ wch: Math.min(maxW + 2, 40) });
        }
        ws['!cols'] = colWidths;

        // Merge school name cell across all columns so it reads as a title
        if (maxCols > 1) {
            if (!ws['!merges']) ws['!merges'] = [];
            ws['!merges'].push({ s:{r:0,c:0}, e:{r:0,c:Math.max(maxCols-1,1)} });
            if (schoolSetup.motto)   ws['!merges'].push({ s:{r:1,c:0}, e:{r:1,c:Math.max(maxCols-1,1)} });
            var examRow = schoolSetup.motto ? (schoolSetup.address ? 3 : 2) : (schoolSetup.address ? 2 : 1);
            ws['!merges'].push({ s:{r:examRow,c:0}, e:{r:examRow,c:Math.max(maxCols-1,1)} });
        }

        XLSX.utils.book_append_sheet(wb, ws, s.name);
    });
    XLSX.writeFile(wb, filename);
}
function downloadClassAnalysis() {
    var exam=getDownloadExam(); if(!exam) return;
    var combos=cleanCombos(exams[exam]||[]);
    var allSubjects=orderedSubjects(combos);

    // Build per-stream per-subject stats
    // Group combos by stream
    var byStream={};
    combos.forEach(function(combo){
        var p=parseCombo(combo);
        if(!byStream[p.stream]) byStream[p.stream]={};
        byStream[p.stream][p.subject]=comboPcts(exam,combo);
    });

    // Header: Subject | Stream | Grade | Learners | Avg% | Highest% | Lowest% | Pass | Fail | Grade
    var rows=[['Subject','Stream','Grade (Stream)','Learners','Avg %','Highest %','Lowest %','Pass (≥50%)','Fail (<50%)','Performance Grade']];

    allSubjects.forEach(function(subj){
        Object.keys(byStream).sort().forEach(function(stream){
            var pcts=byStream[stream][subj]||[];
            var ss=calcStats(pcts); if(!ss) return;
            rows.push([
                subjLabel(subj),
                stream,
                gradeFromStream(stream),
                ss.count,
                +parseFloat(ss.avg.toFixed(2)),
                +ss.highest.toFixed(0),
                +ss.lowest.toFixed(0),
                ss.pass,
                ss.fail,
                isPaperSubject(subj)?(getGrade(ss.avg)||'—'):'—'
            ]);
        });
        rows.push([]); // blank row between subjects
    });

    exportXLSX([{name:'Stream Analysis',data:rows}], exam+'_StreamAnalysis.xlsx', exam);
    showMessage('✓ Stream Analysis downloaded','success');
}
function downloadLevelDistribution() {
    var exam=getDownloadExam(); if(!exam) return;
    if(!grades.length){ showMessage('❌ Define performance levels first','error'); return; }
    var combos=cleanCombos(exams[exam]||[]);
    var allSubjects=orderedSubjects(combos);

    // Merge pcts per logical subject across all streams
    var bySubject={};
    combos.forEach(function(combo){
        var p=parseCombo(combo);
        if(!bySubject[p.subject]) bySubject[p.subject]=[];
        comboPcts(exam,combo).forEach(function(pct){ bySubject[p.subject].push(pct); });
    });

    var hdr=['Subject','Total Learners']
        .concat(grades.map(function(g){ return g.grade+' Count'; }))
        .concat(grades.map(function(g){ return g.grade+' %'; }));
    var rows=[hdr];

    allSubjects.forEach(function(subj){
        var pcts=bySubject[subj]||[]; if(!pcts.length) return;
        var total=pcts.length;
        var counts=grades.map(function(g){ return pcts.filter(function(p){ return p>=g.minScore&&p<=g.maxScore; }).length; });
        var pcents=counts.map(function(c){ return total?+(c/total*100).toFixed(0):0; });
        rows.push([subjLabel(subj),total].concat(counts).concat(pcents));
    });

    exportXLSX([{name:'Level Distribution',data:rows}], exam+'_LevelDistribution.xlsx', exam);
    showMessage('✓ Level Distribution downloaded','success');
}
function buildLearnerReportRows(exam, ranked) {
    var combos = cleanCombos(exams[exam]||[]);
    var allSubjects = orderedSubjects(combos);
    var learnerData = {}, learnerOrder = [];

    combos.forEach(function(combo){
        var p=parseCombo(combo), key=exam+'_'+combo;
        Object.entries(examMarks[key]||{}).forEach(function(e){
            var name=e[0], raw=e[1];
            if(!learnerData[name]){
                learnerData[name]={stream:p.stream,grade:gradeFromStream(p.stream),pct:{}};
                learnerOrder.push(name);
            }
            learnerData[name].pct[p.subject]=toPercent(raw,p.subject,p.stream);
        });
    });

    var entries = learnerOrder.map(function(n){ return [n, learnerData[n]]; });
    if(ranked){
        entries.sort(function(a,b){
            return meanPct(b[1].pct) - meanPct(a[1].pct);
        });
    }

    var hdr = ['#', 'Learner', 'Stream', 'Grade'];
    allSubjects.forEach(function(s){
        hdr.push(subjLabel(s) + ' %');
        hdr.push(subjLabel(s) + ' Level');
    });
    hdr.push('Sum', 'Total %', 'Total Level');

    var rows = [hdr];
    entries.forEach(function(entry, idx){
        var name=entry[0], d=entry[1];
        var sum      = Math.round(sumPcts(d.pct));
        var totalPct = Math.round(meanPct(d.pct));
        var row = [idx+1, name, d.stream, d.grade];
        allSubjects.forEach(function(s){
            var pct = d.pct[s];
            row.push(pct !== undefined ? Math.round(pct) : '');
            row.push(pct !== undefined ? (getGrade(pct) || '—') : '');
        });
        row.push(sum, totalPct, getGrade(totalPct) || '—');
        rows.push(row);
    });
    return rows;
}

function downloadFullReport() {
    var exam=getDownloadExam(); if(!exam) return;
    var lRows=buildLearnerReportRows(exam, true);
    var sRows=buildStreamAnalysisRows(exam);
    exportXLSX([{name:'Ranked Results',data:lRows},{name:'Stream Analysis',data:sRows}], exam+'_FullReport_Ranked.xlsx', exam);
    showMessage('✓ Ranked Full Report downloaded','success');
}

function downloadUnrankedReport() {
    var exam=getDownloadExam(); if(!exam) return;
    var lRows=buildLearnerReportRows(exam, false);
    var sRows=buildStreamAnalysisRows(exam);
    exportXLSX([{name:'Unranked Results',data:lRows},{name:'Stream Analysis',data:sRows}], exam+'_FullReport_Unranked.xlsx', exam);
    showMessage('✓ Unranked Full Report downloaded','success');
}

function buildStreamAnalysisRows(exam) {
    var combos=cleanCombos(exams[exam]||[]);
    var rows=[['Exam','Stream','Gr','Subject','Learners','Avg %','Highest %','Lowest %','Pass (≥50%)','Fail (<50%)','Grade']];
    // order by stream then fixed subject order
    var allSubjects=orderedSubjects(combos);
    var streams=[...new Set(combos.map(function(c){ return parseCombo(c).stream; }))].sort();
    streams.forEach(function(stream){
        allSubjects.forEach(function(subj){
            var key=exam+'_'+stream+'_'+subj;
            if(!examMarks[key]) return;
            var pcts=Object.values(examMarks[key]).map(function(r){ return toPercent(r,subj,st); });
            var ss=calcStats(pcts); if(!ss) return;
            rows.push([exam,stream,gradeFromStream(stream),subjLabel(subj),ss.count,+parseFloat(ss.avg.toFixed(2)),+ss.highest.toFixed(0),+ss.lowest.toFixed(0),ss.pass,ss.fail,isPaperSubject(subj)?(getGrade(ss.avg)||'—'):'—']);
        });
    });
    return rows;
}
function downloadGradeComparison() {
    var exam=getDownloadExam(); if(!exam) return;
    var combos=cleanCombos(exams[exam]||[]);
    var byGrade={};
    combos.forEach(function(combo){
        var p=parseCombo(combo), g=gradeFromStream(p.stream);
        if(!byGrade[g]) byGrade[g]={streams:{},subjects:[]};
        if(!byGrade[g].streams[p.stream]) byGrade[g].streams[p.stream]={};
        if(!byGrade[g].subjects.includes(p.subject)) byGrade[g].subjects.push(p.subject);
        byGrade[g].streams[p.stream][p.subject]=comboPcts(exam,combo);
    });
    var sheets=[];
    Object.keys(byGrade).sort().forEach(function(g){
        var gd=byGrade[g], streamNames=Object.keys(gd.streams).sort();
        var rows=[['Subject','Stream','Learners','Avg %','Highest %','Lowest %','Pass','Fail','Grade']
            .concat(grades.map(function(gr){ return gr.grade+' Count'; }))];
        var ordSubjs=orderedSubjects(gd.subjects);
        ordSubjs.forEach(function(subject){
            streamNames.forEach(function(st){
                var pcts=(gd.streams[st][subject])||[];
                var ss=calcStats(pcts); if(!ss) return;
                var row=[subjLabel(subject),st,ss.count,+parseFloat(ss.avg.toFixed(2)),+ss.highest.toFixed(0),+ss.lowest.toFixed(0),ss.pass,ss.fail,isPaperSubject(subject)?(getGrade(ss.avg)||'—'):'—'];
                grades.forEach(function(gr){ row.push(pcts.filter(function(p){ return p>=gr.minScore&&p<=gr.maxScore; }).length); });
                rows.push(row);
            });
            rows.push([]); // spacer
        });
        sheets.push({name:'Grade '+g,data:rows});
    });
    if(!sheets.length){ showMessage('❌ No data to export','error'); return; }
    exportXLSX(sheets, exam+'_GradeComparison.xlsx', exam);
    showMessage('✓ Grade Comparison downloaded','success');
}

// ── GRADE RANKING: all learners across streams, ranked within each grade ──
function downloadGradeRanking() {
    var exam=getDownloadExam(); if(!exam) return;
    var combos=cleanCombos(exams[exam]||[]);
    if(!combos.length){ showMessage('❌ No data for this exam','error'); return; }
    var allSubjects=orderedSubjects(combos);

    // Collect per-learner subject percentages
    var learnerData={}; // {name: {stream, grade, pct:{subj:pct}}}
    combos.forEach(function(combo){
        var p=parseCombo(combo), key=exam+'_'+combo;
        Object.entries(examMarks[key]||{}).forEach(function(e){
            var name=e[0], raw=e[1];
            if(!learnerData[name]) learnerData[name]={stream:p.stream,grade:gradeFromStream(p.stream),pct:{}};
            learnerData[name].pct[p.subject]=toPercent(raw,p.subject,p.stream);
        });
    });

    // Group learners by grade
    var byGrade={};
    Object.entries(learnerData).forEach(function(e){
        var name=e[0], d=e[1];
        if(!byGrade[d.grade]) byGrade[d.grade]=[];
        byGrade[d.grade].push({name:name,stream:d.stream,pct:d.pct});
    });

    var sheets=[];

    Object.keys(byGrade).sort().forEach(function(g){
        var group=byGrade[g];

        // Sort by meanPct descending
        group.sort(function(a,b){ return meanPct(b.pct)-meanPct(a.pct); });

        // Subjects present in this grade
        var gradeCombos=combos.filter(function(c){ return gradeFromStream(parseCombo(c).stream)===g; });
        var gradeSubjects=orderedSubjects(gradeCombos);

        // Build header
        var hdr=['Rank','Learner','Stream'].concat(
            gradeSubjects.map(function(s){ return subjLabel(s)+' %'; })
        ).concat(['Sum','Total %','Level']);

        var rows=[
            ['Grade '+g+' — Combined Stream Ranking'],
            ['Exam: '+exam],
            ['Streams: '+[...new Set(group.map(function(d){return d.stream;}))].sort().join(', ')],
            ['Total learners: '+group.length],
            [],
            hdr
        ];

        group.forEach(function(d,idx){
            var sum=Math.round(sumPcts(d.pct));
            var total=Math.round(meanPct(d.pct));
            var row=[idx+1, d.name, d.stream];
            gradeSubjects.forEach(function(s){
                var pct=d.pct[s];
                row.push(pct!==undefined ? Math.round(pct)+'%' : '—');
            });
            row.push(sum+'%', total+'%', getGrade(total)||'—');
            rows.push(row);
        });

        sheets.push({name:'Grade '+g, data:rows});
    });

    if(!sheets.length){ showMessage('❌ No data to export','error'); return; }
    exportXLSX(sheets, exam+'_GradeRanking.xlsx', exam);
    showMessage('✓ Grade Ranking downloaded — one sheet per grade','success');
}



function downloadReportCards() {
    // Collect selected exams (1–3)
    var selectedExams = ['rcExam1','rcExam2','rcExam3']
        .map(function(id){ return document.getElementById(id).value; })
        .filter(Boolean);
    if (!selectedExams.length) { showMessage('❌ Select at least one exam','error'); return; }

    var filterLearner = document.getElementById('reportCardLearner').value;
    var filterStream  = document.getElementById('reportCardStream').value;

    // Build union of all subjects across all selected exams
    var allSubjects = orderedSubjects(
        selectedExams.reduce(function(acc,exam){
            return acc.concat(cleanCombos(exams[exam]||[]));
        },[])
    );

    // For each exam collect learnerData: {name: {stream, subjects:{subj:{pct}}, totalPct}}
    function collectExamData(exam) {
        var data = {};
        cleanCombos(exams[exam]||[]).forEach(function(combo){
            var p=parseCombo(combo), key=exam+'_'+combo;
            if(filterStream && p.stream!==filterStream) return;
            Object.entries(examMarks[key]||{}).forEach(function(e){
                var name=e[0], raw=e[1];
                if(!data[name]) data[name]={stream:p.stream,subjects:{},pctMap:{}};
                var pct=toPercent(raw,p.subject,p.stream);
                data[name].subjects[p.subject]=pct;
                data[name].pctMap[p.subject]=pct;
            });
        });
        // totalPct = mean of subject percentages
        Object.values(data).forEach(function(d){
            d.pctMap = d.subjects; // alias for sumPcts
            d.totalPct = Math.round(meanPct(d.pctMap));
        });
        return data;
    }

    var examDatasets = selectedExams.map(function(e){ return collectExamData(e); });

    // Master learner list — union across all exams, respecting stream filter
    var masterNames = new Set();
    examDatasets.forEach(function(ds){ Object.keys(ds).forEach(function(n){ masterNames.add(n); }); });
    if(filterLearner){
        var parts=filterLearner.split('||'), pick=parts[0];
        masterNames = new Set([pick].filter(function(n){ return masterNames.has(n); }));
    }
    var learnerOrder = [...masterNames].sort(function(a,b){
        // Sort by stream of first available dataset then name
        var sa='', sb='';
        examDatasets.forEach(function(ds){ if(ds[a]&&!sa) sa=ds[a].stream; if(ds[b]&&!sb) sb=ds[b].stream; });
        return sa.localeCompare(sb)||a.localeCompare(b);
    });
    if(!learnerOrder.length){ showMessage('❌ No data found','error'); return; }

    var school = schoolSetup.name || 'SCHOOL NAME';
    var numExams = selectedExams.length;

    // B&W-safe level helpers — use distinct patterns/shading + text, not just color
    // Each level gets a unique fill shade so they differ in greyscale print
    function lvlStyle(g) {
        var i = grades.findIndex(function(gr){ return gr.grade===g; });
        // background, text color, border — chosen so greyscale contrast is clear
        if (i===0) return {bg:'#000000',fg:'#ffffff',bd:'#000000'};  // black  = top
        if (i===1) return {bg:'#444444',fg:'#ffffff',bd:'#444444'};  // dark grey
        if (i===2) return {bg:'#888888',fg:'#ffffff',bd:'#888888'};  // mid grey
        if (i===3) return {bg:'#cccccc',fg:'#000000',bd:'#888888'};  // light grey
        return            {bg:'#ffffff',fg:'#000000',bd:'#aaaaaa'};  // white
    }
    function lvlBadge(g) {
        if (!g || g==='—') return '<span style="font-size:11px;color:#000;">—</span>';
        var s=lvlStyle(g);
        return '<span style="display:inline-block;min-width:32px;background:'+s.bg+';color:'+s.fg
            +';font-weight:900;font-size:11px;padding:2px 8px;border-radius:3px;border:1.5px solid '
            +s.bd+';letter-spacing:0.5px;text-align:center;">'+g+'</span>';
    }
    function devArrow(dev) {
        if (dev===null) return '<span style="color:#666;">—</span>';
        if (dev>0)  return '<span style="font-weight:800;font-size:13px;">▲ +'+dev+'%</span>';
        if (dev<0)  return '<span style="font-weight:800;font-size:13px;">▼ '+dev+'%</span>';
        return '<span style="font-weight:700;">→ 0%</span>';
    }

    // Performance key — text + shaded boxes readable in B&W
    var perfKey = grades.length ? grades.map(function(g){
        var s=lvlStyle(g.grade);
        return '<span style="display:inline-block;background:'+s.bg+';color:'+s.fg
            +';border:1.5px solid '+s.bd+';font-weight:800;font-size:10px;padding:2px 9px'
            +';border-radius:3px;margin-right:5px;margin-bottom:3px;">'+g.grade+': '+g.minScore+'–'+g.maxScore+'%</span>';
    }).join('') : '<span style="font-size:11px;color:#444;">No performance levels defined</span>';

    var cards = learnerOrder.map(function(name, cardIdx){
        var stream='', streamGrade='';
        examDatasets.forEach(function(ds){
            if(ds[name]&&!stream){ stream=ds[name].stream; streamGrade=gradeFromStream(stream); }
        });
        var classTeacher = (teachers.streams[stream])||{name:'',sig:''};
        var principal    = teachers.principal||{name:'',title:'Principal',sig:''};

        // Table header style — solid black, white text
        var thS = 'padding:8px 7px;font-size:10px;text-transform:uppercase;letter-spacing:0.4px;color:#fff;font-weight:800;border:1px solid #000;';

        // Header: Subject | Max | [Exam Score % | Level] per exam | [Deviation if ≥2 exams]
        // No Sum/Total columns in subject rows — those appear only in the Total/Overall footer
        var thCells = '<th style="'+thS+'background:#000;text-align:left;width:26%;">Subject</th>'
            +'<th style="'+thS+'background:#000;text-align:center;width:7%;">Max</th>';
        selectedExams.forEach(function(exam){
            thCells += '<th style="'+thS+'background:#222;text-align:center;">'+exam+'<br><span style="font-size:8px;font-weight:400;opacity:0.85;">Assessment Score %</span></th>'
                      +'<th style="'+thS+'background:#555;text-align:center;">Level</th>';
        });
        if (numExams>=2) {
            var lastTwo=selectedExams.slice(-2);
            thCells += '<th style="'+thS+'background:#444;text-align:center;">Deviation<br><span style="font-size:8px;font-weight:400;">'+lastTwo[0]+' → '+lastTwo[1]+'</span></th>';
        }

        // Subject rows — Score % and Level per exam only, no Sum/Total
        var subjectRows = allSubjects.map(function(subj, si){
            var hasAny = examDatasets.some(function(ds){ return ds[name]&&ds[name].subjects[subj]!==undefined; });
            if (!hasAny) return '';
            var isP  = isPaperSubject(subj);
            var maxM = getMaxMarks(subj, stream);
            var rowBg = si%2===0 ? '#ffffff' : '#eeeeee';
            var cellBorder = 'border:1px solid #bbb;';

            var cells = '<td style="padding:7px 8px;'+cellBorder+'font-size:12px;font-weight:700;background:'+rowBg+';">'
                +subjLabel(subj)+(isP?' <span style="font-size:9px;background:#000;color:#fff;padding:1px 4px;border-radius:2px;margin-left:3px;font-weight:700;">P1+P2</span>':'')
                +'</td>'
                +'<td style="padding:7px 6px;'+cellBorder+'text-align:center;font-size:11px;background:'+rowBg+';">'+maxM+'</td>';

            var pctValues=[];
            examDatasets.forEach(function(ds){
                var pct=(ds[name]&&ds[name].subjects[subj]!==undefined) ? Math.round(ds[name].subjects[subj]) : null;
                pctValues.push(pct);
                if (pct===null) {
                    cells += '<td style="padding:7px 6px;'+cellBorder+'text-align:center;background:'+rowBg+';color:#999;">—</td>'
                            +'<td style="padding:7px 6px;'+cellBorder+'text-align:center;background:'+rowBg+';">—</td>';
                } else {
                    var g=getGrade(pct)||'—';
                    cells += '<td style="padding:7px 6px;'+cellBorder+'text-align:center;font-weight:800;font-size:13px;background:'+rowBg+';">'+pct+'%</td>'
                            +'<td style="padding:7px 6px;'+cellBorder+'text-align:center;background:'+rowBg+';">'+lvlBadge(g)+'</td>';
                }
            });
            // Deviation column (subject level)
            if (numExams>=2) {
                var p1=pctValues[numExams-2], p2=pctValues[numExams-1];
                var dev=(p1!==null&&p2!==null)?(p2-p1):null;
                cells += '<td style="padding:7px 6px;'+cellBorder+'text-align:center;font-size:12px;background:'+rowBg+';">'+devArrow(dev)+'</td>';
            }
            return '<tr>'+cells+'</tr>';
        }).join('');

        // Total / Overall footer row
        // Layout: TOTAL/OVERALL (colspan 2) | [Score% placeholder | Level placeholder] per exam | Sum | Avg % | Level | [Deviation]
        var tfS='padding:9px 7px;font-weight:800;font-size:12px;background:#000;color:#fff;border:1px solid #000;';
        var tfGrey='padding:9px 7px;font-weight:800;font-size:12px;background:#333;color:#fff;border:1px solid #000;';

        var totalCells='<td style="'+tfS+'" colspan="2">TOTAL / OVERALL</td>';

        // Blank placeholders under Score% and Level columns (per exam)
        examDatasets.forEach(function(ds){
            totalCells+='<td style="'+tfS+'text-align:center;opacity:0.4;font-size:10px;">—</td>'
                       +'<td style="'+tfS+'text-align:center;opacity:0.4;font-size:10px;">—</td>';
        });

        // Sum | Avg % | Level — one set per exam, shown in grey cells adjacent to placeholder cols
        // Actually: we put Sum then Avg% then Level for EACH exam in the same row at the end
        var totalPcts=[];
        examDatasets.forEach(function(ds, ei){
            var tp=(ds[name])?ds[name].totalPct:null;
            var sm=(ds[name])?Math.round(sumPcts(ds[name].subjects||{})):null;
            totalPcts.push(tp);
        });

        // Remove per-exam placeholders above — instead show each exam's Sum|Avg%|Level in the total row
        // Rebuild: colspan 2 label, then for each exam: score%=blank, level=blank, THEN at the end: Sum | Avg% | Level per exam
        // Simpler: put "Sum: X | Avg: Y% | Level" merged into the remaining columns
        // Best approach given column count: rebuild totalCells fresh
        totalCells = '<td style="'+tfS+'" colspan="2">TOTAL / OVERALL</td>';
        examDatasets.forEach(function(ds, ei){
            var tp=(ds[name])?ds[name].totalPct:null;
            var sm=(ds[name])?Math.round(sumPcts(ds[name].subjects||{})):null;
            totalPcts[ei]=tp;
            if (tp===null){
                totalCells+='<td style="'+tfS+'text-align:center;opacity:0.5;">—</td>'
                           +'<td style="'+tfS+'text-align:center;opacity:0.5;">—</td>';
            } else {
                var tg=getGrade(tp)||'—';
                // Score% col → shows Sum
                totalCells+='<td style="'+tfGrey+'text-align:center;font-size:13px;" title="Sum of all subject %">'+(sm!==null?sm:'—')+'</td>'
                // Level col → shows Avg% + badge
                           +'<td style="'+tfGrey+'text-align:center;font-size:12px;">'+tp+'% '+lvlBadge(tg)+'</td>';
            }
        });
        if (numExams>=2) {
            var tp1=totalPcts[numExams-2], tp2=totalPcts[numExams-1];
            var tdev=(tp1!==null&&tp2!==null)?(tp2-tp1):null;
            totalCells+='<td style="'+tfS+'text-align:center;font-size:13px;">'+devArrow(tdev)+'</td>';
        }

        return '<div class="report-card">'

        // ── HEADER ── thick black top bar
        +'<div style="background:#000;color:#fff;padding:14px 20px;border-radius:0;">'
        +  '<div style="display:flex;align-items:center;gap:14px;">'
        +  (schoolSetup.logo
            ? '<img src="'+schoolSetup.logo+'" style="height:60px;width:60px;object-fit:contain;background:#fff;border-radius:4px;padding:3px;flex-shrink:0;">'
            : '<div style="height:60px;width:60px;background:#444;border-radius:4px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:26px;">🏫</div>'
          )
        +  '<div style="flex:1;text-align:center;">'
        +    '<div style="font-size:17px;font-weight:900;letter-spacing:0.8px;text-transform:uppercase;">'+school+'</div>'
        +    (schoolSetup.motto   ?'<div style="font-size:10px;opacity:0.85;margin-top:2px;font-style:italic;">'+schoolSetup.motto+'</div>':'')
        +    (schoolSetup.address ?'<div style="font-size:10px;opacity:0.75;margin-top:1px;">'+schoolSetup.address+'</div>':'')
        +    '<div style="margin-top:6px;font-size:11px;font-weight:700;border:1.5px solid #fff;display:inline-block;padding:2px 14px;border-radius:2px;letter-spacing:1px;">ACADEMIC PERFORMANCE REPORT FORM</div>'
        +  '</div>'
        +  '<div style="text-align:right;flex-shrink:0;font-size:10px;opacity:0.8;">No. '+(cardIdx+1)+'</div>'
        +  '</div>'
        +'</div>'

        // ── LEARNER INFO ── heavy bordered strip
        +'<div style="border:2px solid #000;border-top:none;padding:10px 20px;background:#f0f0f0;">'
        +  '<div style="display:grid;grid-template-columns:2fr 1fr 1fr;gap:10px;font-size:12px;">'
        +    '<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#444;margin-bottom:2px;">Learner Name</div>'
        +      '<strong style="font-size:15px;color:#000;text-transform:uppercase;">'+name+'</strong></div>'
        +    '<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#444;margin-bottom:2px;">Stream</div>'
        +      '<strong style="font-size:14px;">'+stream+'</strong></div>'
        +    '<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#444;margin-bottom:2px;">Grade / Year</div>'
        +      '<strong style="font-size:14px;">Grade '+streamGrade+'</strong></div>'
        +  '</div>'
        +'</div>'

        // ── MARKS TABLE ──
        +'<div style="padding:12px 20px;">'
        +  '<table style="width:100%;border-collapse:collapse;border:2px solid #000;font-size:12px;">'
        +    '<thead><tr>'+thCells+'</tr></thead>'
        +    '<tbody>'+subjectRows+'</tbody>'
        +    '<tfoot><tr>'+totalCells+'</tr></tfoot>'
        +  '</table>'
        +'</div>'

        // ── PERFORMANCE KEY + REMARKS ──
        +'<div style="padding:0 20px 12px;display:grid;grid-template-columns:1fr 1fr;gap:12px;">'
        +  '<div style="border:1.5px solid #000;border-radius:3px;padding:8px 10px;">'
        +    '<div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;border-bottom:1px solid #aaa;padding-bottom:3px;">Performance Level Key</div>'
        +    '<div>'+perfKey+'</div>'
        +  '</div>'
        +  '<div style="border:1.5px solid #000;border-radius:3px;padding:8px 10px;">'
        +    '<div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;border-bottom:1px solid #aaa;padding-bottom:3px;">Class Teacher\'s Remarks</div>'
        +    '<div style="min-height:38px;border-bottom:1px solid #aaa;margin-bottom:2px;"></div>'
        +  '</div>'
        +'</div>'

        // ── SIGNATURES ──
        +'<div style="padding:0 20px 18px;display:grid;grid-template-columns:1fr 1fr;gap:16px;">'
        +  '<div style="border:1.5px solid #000;border-radius:3px;padding:10px 12px;text-align:center;">'
        +    '<div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Class Teacher</div>'
        +    (classTeacher.sig
              ? '<img src="'+classTeacher.sig+'" style="max-height:40px;max-width:120px;object-fit:contain;display:block;margin:0 auto 5px;filter:grayscale(100%);">'
              : '<div style="height:36px;border-bottom:1.5px solid #000;margin-bottom:5px;"></div>')
        +    '<div style="font-size:11px;font-weight:700;color:#000;">'+(classTeacher.name||'______________________')+'</div>'
        +    '<div style="font-size:9px;color:#444;margin-top:2px;">Class Teacher — '+stream+'</div>'
        +  '</div>'
        +  '<div style="border:1.5px solid #000;border-radius:3px;padding:10px 12px;text-align:center;">'
        +    '<div style="font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">'+(principal.title||'Principal')+'</div>'
        +    (principal.sig
              ? '<img src="'+principal.sig+'" style="max-height:40px;max-width:120px;object-fit:contain;display:block;margin:0 auto 5px;filter:grayscale(100%);">'
              : '<div style="height:36px;border-bottom:1.5px solid #000;margin-bottom:5px;"></div>')
        +    '<div style="font-size:11px;font-weight:700;color:#000;">'+(principal.name||'______________________')+'</div>'
        +    '<div style="font-size:9px;color:#444;margin-top:2px;">'+(principal.title||'Principal')+'</div>'
        +  '</div>'
        +'</div>'

        // ── BOTTOM BAR ── solid black rule
        +'<div style="height:6px;background:#000;border-radius:0;"></div>'
        +'</div>'; // end .report-card
    }).join('\n');

    var examTitle = selectedExams.join(' / ');
    var html = '<!DOCTYPE html><html lang="en"><head>'
        +'<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
        +'<title>'+school+' — '+examTitle+' Report Cards</title>'
        +'<style>'
        +'*{box-sizing:border-box;margin:0;padding:0;}'
        +'body{font-family:Arial,sans-serif;background:#ccc;padding:20px;color:#000;}'
        +'.report-card{background:#fff;border:2px solid #000;max-width:760px;margin:0 auto 32px;page-break-after:always;}'
        +'.report-card:last-child{page-break-after:avoid;}'
        +'@media print{'
        +  'body{background:#fff;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}'
        +  '.no-print{display:none!important;}'
        +  '.report-card{border:1.5px solid #000;margin:0;max-width:100%;page-break-after:always;box-shadow:none;}'
        +  '.report-card:last-child{page-break-after:avoid;}'
        +'}'
        +'</style>'
        +'</head><body>'
        +'<div class="no-print" style="max-width:760px;margin:0 auto 18px;background:#fff;border:1px solid #aaa;padding:12px 18px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
        +  '<button onclick="window.print()" style="background:#000;color:#fff;border:none;padding:9px 22px;font-size:14px;font-weight:700;border-radius:3px;cursor:pointer;">🖨️ Print / Save as PDF</button>'
        +  '<button onclick="window.close()" style="background:#666;color:#fff;border:none;padding:9px 14px;font-size:14px;font-weight:700;border-radius:3px;cursor:pointer;">✕ Close</button>'
        +  '<span style="font-size:12px;color:#444;">'+learnerOrder.length+' card'+(learnerOrder.length!==1?'s':'')+' · '+examTitle+'</span>'
        +'</div>'
        +cards
        +'</body></html>';

    var blob=new Blob([html],{type:'text/html'});
    var url=URL.createObjectURL(blob);
    var win=window.open(url,'_blank');
    if(!win){ var a=document.createElement('a');a.href=url;a.download=examTitle+'_ReportCards.html';a.click(); }
    showMessage('✓ '+learnerOrder.length+' report card'+(learnerOrder.length!==1?'s':'')+' opened','success');
}
function clearAllData() {
    if(!confirm('⚠️ Delete ALL data?')) return;
    if(!confirm('🚨 Final confirmation — delete everything?')) return;
    localStorage.clear(); location.reload();
}

// ═══════════════ TAB 8: IMPORT LEARNERS ═══════════════
var learnersWB = null; // cached workbook

function loadLearnersWorkbook() {
    var file=document.getElementById('learnersFile').files[0];
    document.getElementById('learnersPreview').innerHTML='';
    document.getElementById('learnersSheetPicker').style.display='none';
    learnersWB=null;
    if(!file) return;
    var reader=new FileReader();
    reader.onload=function(e){
        try {
            learnersWB=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
            var sheets=learnersWB.SheetNames;
            var sel=document.getElementById('learnersSheetSelect');
            sel.innerHTML='';
            sheets.forEach(function(s){ var o=document.createElement('option');o.value=s;o.textContent=s;sel.appendChild(o); });
            document.getElementById('learnersSheetPicker').style.display= sheets.length>1 ? 'block' : 'none';
            previewLearners(); // auto-preview first sheet
        } catch(err){ showMessage('❌ Error reading file: '+err.message,'error'); }
    };
    reader.readAsArrayBuffer(file);
}

function previewLearners() {
    if (!learnersWB) { showMessage('❌ Upload a file first', 'error'); return; }
    var sheetName = document.getElementById('learnersSheetSelect').value || learnersWB.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(learnersWB.Sheets[sheetName], { header: 1 });
    var html = '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:4px;padding:8px 12px;margin-bottom:10px;font-size:13px;">'
        + '📄 Sheet: <strong>' + sheetName + '</strong>'
        + (learnersWB.SheetNames.length > 1 ? ' &nbsp;(' + learnersWB.SheetNames.length + ' sheets in workbook)' : '')
        + '<br>Headers from <strong>Row 5</strong>, data from <strong>Row 6</strong> onwards.</div>';
    html += '<div class="tbl-wrap"><table><tr><th>Serial (Col A)</th><th>Name (Col B)</th><th>Gender (Col C)</th><th>Stream (Col D)</th></tr>';
    var total = 0;
    // Data starts at row 6 = index 5
    for (var i = 5; i < rows.length; i++) {
        var r = rows[i]; if (!r || !r[1]) continue; total++;
        if (total <= 10) html += '<tr><td>' + (r[0] || '') + '</td><td>' + String(r[1] || '') + '</td><td>' + String(r[2] || '') + '</td><td>'
            + '<span class="stream-tag">' + String(r[3] || '—') + '</span></td></tr>';
    }
    if (total > 10) html += '<tr><td colspan="4" style="color:#6b7280;font-size:12px;">... and ' + (total - 10) + ' more rows</td></tr>';
    html += '</table></div><p style="margin-top:8px;color:#6b7280;font-size:12px;">Total learner rows: <strong>' + total + '</strong></p>';
    document.getElementById('learnersPreview').innerHTML = html;
    showMessage('✓ Sheet "' + sheetName + '" — ' + total + ' learners found', 'success');
}

function importLearners() {
    if (!learnersWB) { showMessage('❌ Select a file first', 'error'); return; }
    var sheetName = document.getElementById('learnersSheetSelect').value || learnersWB.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(learnersWB.Sheets[sheetName], { header: 1 });
    var added = 0, clsAdded = new Set();
    // Data starts at row 6 = index 5
    for (var i = 5; i < rows.length; i++) {
        var r = rows[i]; if (!r || !r[1]) continue;
        var name   = String(r[1] || '').trim();
        var gender = String(r[2] || '').trim();
        var stream = String(r[3] || '').trim();
        if (!name || !stream) continue;
        if (!classes.includes(stream)) { classes.push(stream); clsAdded.add(stream); }
        if (!learners[stream]) learners[stream] = [];
        if (!learners[stream].find(function(l) { return l.name === name; })) {
            learners[stream].push({ name: name, gender: gender }); added++;
        }
    }
    saveAllData();
    document.getElementById('learnersFile').value = '';
    document.getElementById('learnersPreview').innerHTML = '';
    document.getElementById('learnersSheetPicker').style.display = 'none';
    learnersWB = null;
    refreshAllDropdowns();
    showMessage('✓ Imported ' + added + ' learners across ' + clsAdded.size + ' new streams — opening Learners Store…', 'success');
    setTimeout(function() { switchTab(10, null); }, 800);
}

// ═══════════════ TAB 9: IMPORT MARKS ═══════════════
var marksWB = null; // cached workbook

function loadMarksWorkbook() {
    var file=document.getElementById('marksFile').files[0];
    document.getElementById('marksPreview').innerHTML='';
    document.getElementById('marksSheetPicker').style.display='none';
    marksWB=null;
    if(!file) return;
    var reader=new FileReader();
    reader.onload=function(e){
        try {
            marksWB=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
            var sheets=marksWB.SheetNames;
            var sel=document.getElementById('marksSheetSelect');
            sel.innerHTML='';
            sheets.forEach(function(s){ var o=document.createElement('option');o.value=s;o.textContent=s;sel.appendChild(o); });
            document.getElementById('marksSheetPicker').style.display= sheets.length>1 ? 'block' : 'none';
            previewMarks(); // auto-preview first sheet
        } catch(err){ showMessage('❌ Error reading file: '+err.message,'error'); }
    };
    reader.readAsArrayBuffer(file);
}

// ── Parse subject headers from Row 5 (index 4), col E onwards (index 4+) ──
// Rules:
//   - Blank cells → separator, skip entirely
//   - ENGP1, ENG P1, ENG1, ENGLISH P1 etc → logicalName = 'English Total' (summed with P2)
//   - ENGP2, ENG P2, ENG2, ENGLISH P2 etc → logicalName = 'English Total' (added to P1)
//   - KISP1, KIS P1, KIS1, KISWAHILI P1 etc → logicalName = 'Kiswahili Total' (summed with P2)
//   - KISP2, KIS P2, KIS2, KISWAHILI P2 etc → logicalName = 'Kiswahili Total' (added to P1)
//   - Any other non-blank header → logicalName = header as-is
// Returns [{colIdx, rawHeader, logicalName}]
function parseMarkHeaders(rows) {
    var headerRow = rows[4] || []; // Row 5 = index 4
    var cols = [];
    for (var c = 4; c < headerRow.length; c++) { // col E = index 4
        var h = String(headerRow[c] || '').trim();
        if (!h) continue; // blank = separator, skip
        var up = h.toUpperCase().replace(/\s+/g, ''); // remove all spaces for matching

        var logical;
        if (up === 'ENGP1' || up === 'ENG1' || up === 'ENGLISHP1' || up === 'ENGLISHPAPER1' || up === 'EP1') {
            logical = 'English Total';
        } else if (up === 'ENGP2' || up === 'ENG2' || up === 'ENGLISHP2' || up === 'ENGLISHPAPER2' || up === 'EP2') {
            logical = 'English Total';
        } else if (up === 'KISP1' || up === 'KIS1' || up === 'KISWAHILIP1' || up === 'KISWAHILIPAPER1' || up === 'KP1') {
            logical = 'Kiswahili Total';
        } else if (up === 'KISP2' || up === 'KIS2' || up === 'KISWAHILIP2' || up === 'KISWAHILIPAPER2' || up === 'KP2') {
            logical = 'Kiswahili Total';
        } else {
            logical = h; // all other subjects kept as-is
        }

        cols.push({ colIdx: c, rawHeader: h, logicalName: logical });
    }
    return cols;
}

// For a data row compute {logicalName: summedRaw}
// English P1 + P2 are summed into 'English Total'
// Kiswahili P1 + P2 are summed into 'Kiswahili Total'
// All other subjects stored as individual values
function computeRowMarks(r, colMap) {
    var result = {};
    colMap.forEach(function(col) {
        var raw = parseFloat(r[col.colIdx]);
        if (isNaN(raw)) return;
        if (result[col.logicalName] === undefined) result[col.logicalName] = 0;
        result[col.logicalName] += raw; // addition handles both single and multi-paper
    });
    return result;
}

function previewMarks() {
    if (!marksWB) { showMessage('❌ Upload a file first', 'error'); return; }
    var sheetName = document.getElementById('marksSheetSelect').value || marksWB.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(marksWB.Sheets[sheetName], { header: 1 });
    if (rows.length < 5) {
        document.getElementById('marksPreview').innerHTML = '<p style="color:#dc2626;">Sheet must have at least 5 rows (headers in row 5).</p>';
        return;
    }

    var colMap = parseMarkHeaders(rows);
    var logicalSubjects = [...new Set(colMap.map(function(c) { return c.logicalName; }))];
    var rawHeaders = colMap.map(function(c) { return c.rawHeader; });

    var html = '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:4px;padding:8px 12px;margin-bottom:10px;font-size:13px;">'
        + '📄 Sheet: <strong>' + sheetName + '</strong>'
        + (marksWB.SheetNames.length > 1 ? ' &nbsp;(' + marksWB.SheetNames.length + ' sheets in workbook)' : '') + '<br>'
        + 'Headers read from: <strong>Row 5, Column E onwards</strong><br>'
        + 'Raw columns found: <strong>' + (rawHeaders.length ? rawHeaders.join(', ') : 'none') + '</strong><br>'
        + 'Subjects after merging papers: <strong>' + (logicalSubjects.length ? logicalSubjects.join(', ') : 'none') + '</strong></div>';

    if (!colMap.length) {
        document.getElementById('marksPreview').innerHTML = html
            + '<p style="color:#dc2626;font-size:13px;">⚠️ No subject headers found in Row 5 from column E onwards. Check your Excel layout.</p>';
        return;
    }

    // Table header — show logical subjects only (English Total / Kiswahili Total, not P1/P2)
    html += '<div class="tbl-wrap"><table><tr><th>#</th><th>Name (Col B)</th><th>Stream (Col D)</th>';
    logicalSubjects.forEach(function(s) {
        var isP = (s === 'English Total' || s === 'Kiswahili Total');
        html += '<th>' + (isP ? '<span style="color:#15803d;font-weight:bold;">' + s + ' ✓</span>' : s) + '</th><th>%</th>';
    });
    html += '</tr>';

    // Data rows start at row 6 (index 5)
    var count = 0;
    for (var i = 5; i < rows.length; i++) {
        var r = rows[i];
        if (!r || !r[1]) continue;
        count++;
        if (count > 10) continue;
        var name   = String(r[1] || '').trim();
        var stream = String(r[3] || '').trim();
        var rowMarks = computeRowMarks(r, colMap);
        html += '<tr><td>' + (r[0] || i) + '</td><td>' + name + '</td><td><span class="stream-tag">' + stream + '</span></td>';
        logicalSubjects.forEach(function(s) {
            var val = rowMarks[s];
            if (val === undefined) {
                html += '<td style="color:#d1d5db;">—</td><td>—</td>';
            } else {
                var pct = toPercent(val,s,stream);
                html += '<td><strong>' + val + '</strong></td><td style="color:#1e40af;font-weight:bold;">' + pct.toFixed(0) + '%</td>';
            }
        });
        html += '</tr>';
    }
    if (count > 10) html += '<tr><td colspan="' + (3 + logicalSubjects.length * 2) + '" style="color:#6b7280;font-size:12px;">... and ' + (count - 10) + ' more rows</td></tr>';
    html += '</table></div><p style="margin-top:8px;color:#6b7280;font-size:12px;">Data rows found: <strong>' + count + '</strong></p>';
    document.getElementById('marksPreview').innerHTML = html;
    showMessage('✓ Sheet "' + sheetName + '" — ' + count + ' learners, ' + logicalSubjects.length + ' subjects detected', 'success');
}

function importMarks() {
    var examName = document.getElementById('importExamName').value;
    if (!examName) { showMessage('❌ Select an active exam from the dropdown — create one in the Exams tab first', 'error'); return; }
    if (!marksWB)  { showMessage('❌ Select a file first', 'error'); return; }
    var sheetName = document.getElementById('marksSheetSelect').value || marksWB.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(marksWB.Sheets[sheetName], { header: 1 });
    if (rows.length < 5) { showMessage('❌ Sheet must have at least 5 rows (headers expected in row 5)', 'error'); return; }

    var colMap = parseMarkHeaders(rows);
    if (!colMap.length) { showMessage('❌ No subject headers found in Row 5 from column E onwards', 'error'); return; }

    if (!exams[examName]) exams[examName] = [];
    var totalMarks = 0, streamsFound = new Set();

    // Data starts at row 6 (index 5)
    for (var i = 5; i < rows.length; i++) {
        var r = rows[i];
        if (!r || !r[1]) continue;
        var name   = String(r[1] || '').trim();
        var gender = String(r[2] || '').trim();
        var stream = String(r[3] || '').trim();
        if (!name || !stream) continue;

        streamsFound.add(stream);
        if (!classes.includes(stream)) classes.push(stream);
        if (!learners[stream]) learners[stream] = [];
        if (!learners[stream].find(function(l) { return l.name === name; }))
            learners[stream].push({ name: name, gender: gender });

        // Summed marks per logical subject (English Total = P1+P2, Kiswahili Total = P1+P2)
        var rowMarks = computeRowMarks(r, colMap);
        Object.entries(rowMarks).forEach(function(e) {
            var logicalSubj = e[0], summedRaw = e[1];
            var key = examName + '_' + stream + '_' + logicalSubj;
            if (!examMarks[key]) examMarks[key] = {};
            examMarks[key][name] = summedRaw;
            totalMarks++;
            var combo = stream + '_' + logicalSubj;
            if (!exams[examName].includes(combo)) exams[examName].push(combo);
        });
    }
    saveAllData();
    document.getElementById('marksFile').value = '';
    // Keep exam selected so user can import more streams for same exam
    document.getElementById('marksPreview').innerHTML = '';
    document.getElementById('marksSheetPicker').style.display = 'none';
    marksWB = null;
    refreshAllDropdowns();
    onImportExamChange(); // refresh info line
    renderExamRecordsList && renderExamRecordsList();
    showMessage('✓ Imported ' + totalMarks + ' subject totals across ' + streamsFound.size + ' streams for "' + examName + '" — opening Analysis…', 'success');
    setTimeout(function() { switchTab(5, null); }, 800);
}

// ═══════════════ TAB 10: LEARNERS STORE ═══════════════
function openLearnersStore() {
    // populate stream filter
    var sel=document.getElementById('learnersStoreStream');
    var cur=sel.value;
    sel.innerHTML='<option value="">-- All Streams --</option>';
    // group by grade for nice ordering
    var sorted=[].concat(classes).sort(function(a,b){
        var ga=gradeFromStream(a), gb=gradeFromStream(b);
        return ga!==gb ? ga.localeCompare(gb) : a.localeCompare(b);
    });
    sorted.forEach(function(c){
        var o=document.createElement('option'); o.value=c; o.textContent=c;
        if(c===cur) o.selected=true;
        sel.appendChild(o);
    });
    renderLearnersStore();
}

function renderLearnersStore() {
    var filterStream=document.getElementById('learnersStoreStream').value;
    var streams=filterStream ? [filterStream] : classes.slice().sort(function(a,b){
        return gradeFromStream(a).localeCompare(gradeFromStream(b))||a.localeCompare(b);
    });
    var container=document.getElementById('learnersStoreContent');
    if(!streams.length){ container.innerHTML='<p style="color:#6b7280;font-size:13px;">No learners imported yet.</p>'; return; }

    var totalShown=0, html='';
    streams.forEach(function(stream){
        var list=learners[stream]||[];
        if(!list.length) return;
        totalShown+=list.length;
        html+='<div class="block" style="margin-bottom:14px;">'
            +'<div class="block-hdr"><span><span class="stream-tag">'+stream+'</span> &nbsp; '+list.length+' learner'+(list.length!==1?'s':'')+'</span>'
            +'<button class="item-delete" onclick="deleteStreamLearners(\''+stream+'\')">🗑️ Delete Stream</button></div>'
            +'<div class="tbl-wrap"><table><tr><th>#</th><th>Name</th><th>Gender</th><th>Action</th></tr>';
        list.forEach(function(l,idx){
            html+='<tr><td>'+(idx+1)+'</td><td><strong>'+l.name+'</strong></td><td>'+l.gender+'</td>'
                +'<td><button class="item-delete" onclick="deleteSingleLearner(\''+stream+'\','+idx+')">Delete</button></td></tr>';
        });
        html+='</table></div></div>';
    });
    if(!totalShown){ container.innerHTML='<p style="color:#6b7280;font-size:13px;">No learners in selected stream.</p>'; return; }
    html='<p style="font-size:12px;color:#6b7280;margin-bottom:10px;">Showing <strong>'+totalShown+'</strong> learner'+(totalShown!==1?'s':'')+' across <strong>'+streams.filter(function(s){ return (learners[s]||[]).length; }).length+'</strong> stream(s).</p>'+html;
    container.innerHTML=html;
}

function deleteSingleLearner(stream, idx) {
    var l=learners[stream]&&learners[stream][idx];
    if(!l) return;
    if(!confirm('Delete "'+l.name+'" from '+stream+'?')) return;
    learners[stream].splice(idx,1);
    saveAllData(); renderLearnersStore();
    showMessage('✓ Learner deleted','success');
}

function deleteStreamLearners(stream) {
    var cnt=(learners[stream]||[]).length;
    if(!confirm('Delete all '+cnt+' learners from stream '+stream+'?')) return;
    learners[stream]=[];
    saveAllData(); renderLearnersStore();
    showMessage('✓ All learners in '+stream+' deleted','success');
}

function deleteAllLearnersInView() {
    var filterStream=document.getElementById('learnersStoreStream').value;
    var streams=filterStream ? [filterStream] : classes.slice();
    var total=streams.reduce(function(n,s){ return n+(learners[s]||[]).length; },0);
    if(!total){ showMessage('No learners to delete','info'); return; }
    if(!confirm('Delete ALL '+total+' learners'+(filterStream?' in '+filterStream:'')+' from the store? This cannot be undone.')) return;
    streams.forEach(function(s){ learners[s]=[]; });
    saveAllData(); renderLearnersStore();
    showMessage('✓ Deleted '+total+' learners','success');
}

// ═══════════════ TAB 11: MARKS STORE ═══════════════
var _marksStoreTab = 'view';

function switchMarksStoreTab(mode) {
    _marksStoreTab = mode;
    document.getElementById('msTabView').classList.toggle('active', mode==='view');
    document.getElementById('msTabEdit').classList.toggle('active', mode==='edit');
    document.getElementById('msViewPane').style.display  = mode==='view'  ? 'block' : 'none';
    document.getElementById('msEditPane').style.display  = mode==='edit'  ? 'block' : 'none';
    if (mode==='edit') renderMarksEdit();
}

function openMarksStore() {
    var examSel=document.getElementById('marksStoreExam');
    var curExam=examSel.value;
    examSel.innerHTML='<option value="">-- Select Exam --</option>';
    Object.keys(exams).forEach(function(e){
        var o=document.createElement('option'); o.value=e; o.textContent=e;
        if(e===curExam) o.selected=true;
        examSel.appendChild(o);
    });
    refreshMarksStoreStreams();
    renderMarksStore();
    if (_marksStoreTab==='edit') renderMarksEdit();
}

function refreshMarksStoreStreams() {
    var examName=document.getElementById('marksStoreExam').value;
    var streamSel=document.getElementById('marksStoreStream');
    var curStream=streamSel.value;
    streamSel.innerHTML='<option value="">-- All Streams --</option>';
    if(!examName) return;
    var streams=new Set((exams[examName]||[]).map(function(c){ return parseCombo(c).stream; }));
    [...streams].sort().forEach(function(s){
        var o=document.createElement('option'); o.value=s; o.textContent=s;
        if(s===curStream) o.selected=true;
        streamSel.appendChild(o);
    });
}

function renderMarksStore() {
    refreshMarksStoreStreams();
    var examName=document.getElementById('marksStoreExam').value;
    var filterStream=document.getElementById('marksStoreStream').value;
    var container=document.getElementById('marksStoreContent');
    if(!examName){ container.innerHTML='<p style="color:#6b7280;font-size:13px;">Select an exam above.</p>'; return; }
    var combos=(exams[examName]||[]).filter(function(c){
        return !filterStream||parseCombo(c).stream===filterStream;
    });
    if(!combos.length){ container.innerHTML='<p style="color:#6b7280;font-size:13px;">No marks for this selection.</p>'; return; }

    var byStream={};
    combos.forEach(function(combo){
        var p=parseCombo(combo);
        if(!byStream[p.stream]) byStream[p.stream]=[];
        byStream[p.stream].push(p.subject);
    });
    var allSubjects=orderedSubjects(combos);
    var totalEntries=0, html='';

    Object.keys(byStream).sort().forEach(function(stream){
        var streamSubjects=byStream[stream];
        var learnerRows={};
        streamSubjects.forEach(function(subj){
            var key=examName+'_'+stream+'_'+subj;
            Object.entries(examMarks[key]||{}).forEach(function(e){
                if(!learnerRows[e[0]]) learnerRows[e[0]]={};
                learnerRows[e[0]][subj]=e[1];
            });
        });
        var names=Object.keys(learnerRows).sort();
        totalEntries+=names.length;

        html+='<div class="block" style="margin-bottom:14px;">'
            +'<div class="block-hdr">'
            +'<span><span class="stream-tag">'+stream+'</span> &nbsp;'+names.length+' learner'+(names.length!==1?'s':'')+' &nbsp;|&nbsp; '+streamSubjects.map(subjLabel).join(', ')+'</span>'
            +'<button class="item-delete" onclick="deleteStreamMarks(\''+examName+'\',\''+stream+'\')">🗑️ Delete Stream</button>'
            +'</div>'
            +'<div class="tbl-wrap"><table><tr><th>#</th><th>Learner</th>';
        allSubjects.forEach(function(s){
            var maxM=getMaxMarks(s,gradeFromStream(stream));
            html+='<th>'+subjLabel(s)+' /'+maxM+'</th><th>%</th>';
        });
        html+='<th></th></tr>';

        names.forEach(function(name,idx){
            html+='<tr><td>'+(idx+1)+'</td><td><strong>'+name+'</strong></td>';
            allSubjects.forEach(function(subj){
                var raw=learnerRows[name][subj];
                if(raw===undefined){ html+='<td style="color:#d1d5db;">—</td><td style="color:#d1d5db;">—</td>'; }
                else {
                    var pct=toPercent(raw,subj,stream);
                    html+='<td><strong>'+raw+'</strong></td>'
                        +'<td style="color:#1e40af;font-weight:bold;">'+Math.round(pct)+'% '+gradeBadge(getGrade(pct))+'</td>';
                }
            });
            var sn=name.replace(/'/g,"\\'");
            html+='<td><button class="item-delete" onclick="deleteSingleMark(\''+examName+'\',\''+stream+'\',\''+sn+'\')">Delete</button></td></tr>';
        });
        html+='</table></div></div>';
    });

    html='<p style="font-size:12px;color:#6b7280;margin-bottom:10px;">Exam: <strong>'+examName+'</strong> &nbsp;|&nbsp; <strong>'+totalEntries+'</strong> entries.</p>'+html;
    container.innerHTML=html;
    if(_marksStoreTab==='edit') renderMarksEdit();
}

// ── EDIT PANE ─────────────────────────────────────────
function renderMarksEdit() {
    var examName    = document.getElementById('marksStoreExam').value;
    var filterStream= document.getElementById('marksStoreStream').value;
    var container   = document.getElementById('marksEditContent');
    if(!container) return;

    if(!examName){
        container.innerHTML='<p style="color:#6b7280;font-size:13px;">Select an exam above.</p>'; return;
    }
    if(!filterStream){
        container.innerHTML='<p style="color:#f59e0b;font-size:13px;font-weight:600;">⚠ Please select a specific stream from the dropdown above to edit marks.</p>'; return;
    }

    var stream  = filterStream;
    var grade   = gradeFromStream(stream);
    var combos  = (exams[examName]||[]).filter(function(c){ return parseCombo(c).stream===stream; });
    if(!combos.length){
        container.innerHTML='<p style="color:#6b7280;font-size:13px;">No marks found for <strong>'+stream+'</strong> in <strong>'+examName+'</strong>.</p>'; return;
    }

    var allSubjects = orderedSubjects(combos);

    // Build learner → {subj: raw} map — also include stream learners with no marks
    var learnerRows = {};
    combos.forEach(function(combo){
        var p=parseCombo(combo), key=examName+'_'+combo;
        Object.entries(examMarks[key]||{}).forEach(function(e){
            if(!learnerRows[e[0]]) learnerRows[e[0]]={};
            learnerRows[e[0]][p.subject]=e[1];
        });
    });
    (learners[stream]||[]).forEach(function(l){
        if(!learnerRows[l.name]) learnerRows[l.name]={};
    });
    var names=Object.keys(learnerRows).sort();

    // Store context on container for save functions
    container.dataset.exam     = examName;
    container.dataset.stream   = stream;
    container.dataset.grade    = grade;
    container.dataset.subjects = JSON.stringify(allSubjects);
    container.dataset.names    = JSON.stringify(names);

    // Action bar
    var html='<div style="margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">'
        +'<button class="btn btn-green" onclick="saveMarksEdits()">💾 Save Edits</button>'
        +'<button class="btn" style="background:#7c3aed;" onclick="saveMarksEditsAndAnalyse()">📊 Save &amp; Analyse</button>'
        +'<span id="editSaveStatus" style="font-size:12px;color:#15803d;font-weight:600;"></span>'
        +'</div>';

    // Table
    html+='<div class="tbl-wrap"><table style="font-size:13px;border-collapse:collapse;">'
        +'<thead><tr>'
        +'<th style="padding:9px 10px;background:#1e3a8a;color:white;white-space:nowrap;">#</th>'
        +'<th style="padding:9px 10px;background:#1e3a8a;color:white;white-space:nowrap;min-width:150px;">Learner</th>';
    allSubjects.forEach(function(s){
        var maxM=getMaxMarks(s,grade);
        var isP=isPaperSubject(s);
        html+='<th style="padding:9px 10px;background:'+(isP?'#166534':'#1e3a8a')+';color:white;text-align:center;white-space:nowrap;">'
            +subjLabel(s)+'<br><span style="font-size:10px;opacity:0.8;">/'+maxM+(isP?' P1+P2':'')+'</span></th>';
    });
    html+='<th style="padding:9px 10px;background:#374151;color:white;white-space:nowrap;min-width:130px;">Live Preview</th>'
        +'</tr></thead><tbody>';

    names.forEach(function(name,idx){
        var rowData=learnerRows[name];
        var odd=idx%2===0;
        html+='<tr style="background:'+(odd?'#ffffff':'#f8fafc')+'">'
            +'<td style="padding:8px 10px;color:#9ca3af;border-bottom:1px solid #e5e7eb;">'+(idx+1)+'</td>'
            +'<td style="padding:8px 10px;font-weight:600;border-bottom:1px solid #e5e7eb;">'+name+'</td>';
        allSubjects.forEach(function(subj,si){
            var raw=rowData[subj];
            var maxM=getMaxMarks(subj,grade);
            var isP=isPaperSubject(subj);
            html+='<td style="padding:6px 8px;border-bottom:1px solid #e5e7eb;text-align:center;background:'+(isP?'#f0fdf4':'inherit')+'">'
                +'<input type="number"'
                +' id="ed_'+idx+'_'+si+'"'
                +' value="'+(raw!==undefined?raw:'')+'"'
                +' placeholder="—" min="0" max="'+maxM+'"'
                +' oninput="liveEditPreview('+idx+')"'
                +' style="width:68px;padding:5px 4px;border:1px solid #d1d5db;border-radius:3px;font-size:13px;text-align:center;">'
                +'</td>';
        });
        html+='<td id="edpv_'+idx+'" style="padding:8px 10px;border-bottom:1px solid #e5e7eb;font-size:12px;"></td>'
            +'</tr>';
    });
    html+='</tbody></table></div>';

    // Footer save bar
    html+='<div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">'
        +'<button class="btn btn-green" onclick="saveMarksEdits()">💾 Save Edits</button>'
        +'<button class="btn" style="background:#7c3aed;" onclick="saveMarksEditsAndAnalyse()">📊 Save &amp; Analyse</button>'
        +'</div>';

    container.innerHTML=html;

    // Initial live previews
    names.forEach(function(_,idx){ liveEditPreview(idx); });
}

function liveEditPreview(idx) {
    var container   = document.getElementById('marksEditContent');
    var allSubjects = JSON.parse(container.dataset.subjects||'[]');
    var stream      = container.dataset.stream||'';
    var cell        = document.getElementById('edpv_'+idx);
    if(!cell) return;

    var pcts=[];
    allSubjects.forEach(function(subj,si){
        var inp=document.getElementById('ed_'+idx+'_'+si);
        if(!inp||inp.value==='') return;
        var raw=parseFloat(inp.value);
        if(!isNaN(raw)) pcts.push(toPercent(raw,subj,stream));
    });

    if(!pcts.length){ cell.innerHTML='<span style="color:#9ca3af;">—</span>'; return; }
    var sum=pcts.reduce(function(a,b){return a+b;},0);
    var avg=Math.round(sum/pcts.length);
    var g=getGrade(avg)||'—';
    cell.innerHTML='<span style="color:#6b7280;">Σ<strong>'+Math.round(sum)+'</strong></span>'
        +' &nbsp; <strong style="color:#1e40af;">'+avg+'%</strong>'
        +' '+gradeBadge(g);
}

function saveMarksEdits() {
    var container   = document.getElementById('marksEditContent');
    if(!container||!container.dataset.exam){
        showMessage('❌ Open Edit tab first','error'); return false;
    }
    var examName    = container.dataset.exam;
    var stream      = container.dataset.stream;
    var allSubjects = JSON.parse(container.dataset.subjects||'[]');
    var names       = JSON.parse(container.dataset.names||'[]');

    var saved=0, cleared=0;
    names.forEach(function(name,idx){
        allSubjects.forEach(function(subj,si){
            var inp=document.getElementById('ed_'+idx+'_'+si);
            if(!inp) return;
            var key=examName+'_'+stream+'_'+subj;
            var combo=stream+'_'+subj;
            if(!exams[examName]) exams[examName]=[];

            if(inp.value===''||inp.value===undefined){
                // Clear mark
                if(examMarks[key]&&examMarks[key][name]!==undefined){
                    delete examMarks[key][name]; cleared++;
                }
            } else {
                var raw=parseFloat(inp.value);
                if(isNaN(raw)) return;
                if(!examMarks[key]) examMarks[key]={};
                examMarks[key][name]=raw; saved++;
                if(!exams[examName].includes(combo)) exams[examName].push(combo);
            }
        });
    });

    saveAllData();
    // Refresh view pane silently
    renderMarksStore();

    var st=document.getElementById('editSaveStatus');
    if(st){
        st.textContent='✓ Saved '+saved+' marks'+(cleared?' · '+cleared+' cleared':'');
        setTimeout(function(){ st.textContent=''; },3500);
    }
    showMessage('✓ Marks updated: '+saved+' saved'+(cleared?', '+cleared+' cleared':''),'success');
    return true;
}

function saveMarksEditsAndAnalyse() {
    if(!saveMarksEdits()) return;
    var examName=document.getElementById('marksEditContent').dataset.exam;
    refreshAllDropdowns();
    var sel=document.getElementById('analysisExam');
    if(sel&&examName) sel.value=examName;
    switchTab(5,null);
    setTimeout(generateAnalysis,250);
}

function deleteSingleMark(examName, stream, learnerName) {
    if(!confirm('Delete all marks for "'+learnerName+'" in '+stream+' ('+examName+')?')) return;
    var combos=(exams[examName]||[]).filter(function(c){ return parseCombo(c).stream===stream; });
    combos.forEach(function(combo){
        var p=parseCombo(combo), key=examName+'_'+combo;
        if(examMarks[key]) delete examMarks[key][learnerName];
    });
    saveAllData(); renderMarksStore();
    showMessage('✓ Marks deleted for '+learnerName,'success');
}

function deleteStreamMarks(examName, stream) {
    if(!confirm('Delete ALL marks for stream '+stream+' in exam "'+examName+'"?')) return;
    var combos=(exams[examName]||[]).filter(function(c){ return parseCombo(c).stream===stream; });
    combos.forEach(function(combo){
        delete examMarks[examName+'_'+combo];
        exams[examName]=exams[examName].filter(function(c){ return c!==combo; });
    });
    saveAllData(); openMarksStore();
    showMessage('✓ All marks for '+stream+' deleted','success');
}

function deleteAllMarksInView() {
    var examName=document.getElementById('marksStoreExam').value;
    var filterStream=document.getElementById('marksStoreStream').value;
    if(!examName){ showMessage('❌ Select an exam first','error'); return; }
    var msg=filterStream
        ? 'Delete ALL marks for stream '+filterStream+' in exam "'+examName+'"?'
        : 'Delete ALL marks for exam "'+examName+'"? This removes the entire exam.';
    if(!confirm(msg)) return;
    if(filterStream){
        deleteStreamMarks(examName,filterStream);
    } else {
        (exams[examName]||[]).forEach(function(combo){ delete examMarks[examName+'_'+combo]; });
        delete exams[examName];
        saveAllData(); openMarksStore();
        showMessage('✓ Entire exam "'+examName+'" deleted','success');
    }
}

// ═══════════════ TAB 13: EXAM MANAGEMENT ═══════════════
function openExamsTab() {
    // pre-fill year
    var yearInput = document.getElementById('newExamYear');
    if (!yearInput.value) yearInput.value = new Date().getFullYear();
    renderExamRecordsList();
}

function createExamRecord() {
    var name  = document.getElementById('newExamName').value.trim();
    var type  = document.getElementById('newExamType').value;
    var year  = document.getElementById('newExamYear').value.trim();
    var term  = document.getElementById('newExamTerm').value;
    if (!name) { showMessage('❌ Enter an exam name','error'); return; }

    // Build full label e.g. "Term 1 2025 - Mid Term"
    var fullName = name;

    // Check for duplicate active exam names
    if (examRecords.find(function(r){ return r.name===fullName && r.status==='active'; })) {
        showMessage('❌ An active exam with this name already exists','error'); return;
    }

    var record = {
        id: Date.now(),
        name: fullName,
        type: type,
        year: year,
        term: term,
        status: 'active',
        createdAt: new Date().toLocaleDateString()
    };
    examRecords.push(record);
    saveAllData();
    refreshAllDropdowns();

    document.getElementById('newExamName').value = '';
    document.getElementById('createExamStatus').textContent = '✓ Created: ' + fullName;
    setTimeout(function(){ document.getElementById('createExamStatus').textContent=''; }, 3000);
    renderExamRecordsList();
    showMessage('✓ Exam created: ' + fullName, 'success');
}

function renderExamRecordsList() {
    var container = document.getElementById('examRecordsList');
    if (!examRecords.length) {
        container.innerHTML = '<p style="padding:14px;color:#6b7280;font-size:13px;">No exams created yet. Use the form above to create your first exam.</p>';
        return;
    }

    // Group by status
    var active = examRecords.filter(function(r){ return r.status==='active'; });
    var closed = examRecords.filter(function(r){ return r.status==='closed'; });

    function examRow(r) {
        var hasMarks = exams[r.name] && exams[r.name].length > 0;
        var markCount = 0;
        if (hasMarks) (exams[r.name]||[]).forEach(function(c){ markCount += Object.keys(examMarks[r.name+'_'+c]||{}).length; });
        return '<tr>'
            + '<td style="padding:10px 12px;">'
            +   '<strong>'+r.name+'</strong>'
            +   '<div style="font-size:11px;color:#6b7280;margin-top:2px;">'+r.type+' &nbsp;·&nbsp; '+r.term+' '+r.year+'</div>'
            + '</td>'
            + '<td style="padding:10px 12px;font-size:12px;color:#6b7280;">'+r.createdAt+'</td>'
            + '<td style="padding:10px 12px;text-align:center;">'
            +   (hasMarks ? '<span style="background:#dbeafe;color:#1e40af;font-size:11px;font-weight:600;padding:2px 8px;border-radius:8px;">'+markCount+' entries</span>'
                          : '<span style="color:#9ca3af;font-size:11px;">No marks</span>')
            + '</td>'
            + '<td style="padding:10px 12px;text-align:center;">'
            +   (r.status==='active'
                ? '<span style="background:#dcfce7;color:#15803d;font-size:11px;font-weight:700;padding:3px 10px;border-radius:10px;">● Active</span>'
                : '<span style="background:#f3f4f6;color:#6b7280;font-size:11px;font-weight:700;padding:3px 10px;border-radius:10px;">Closed</span>')
            + '</td>'
            + '<td style="padding:10px 12px;white-space:nowrap;">'
            +   (r.status==='active'
                ? '<button class="btn" style="background:#d97706;font-size:11px;padding:5px 10px;margin:0 4px 0 0;" onclick="toggleExamStatus('+r.id+')">🔒 Close</button>'
                : '<button class="btn btn-green" style="font-size:11px;padding:5px 10px;margin:0 4px 0 0;" onclick="toggleExamStatus('+r.id+')">🔓 Re-open</button>')
            +   '<button class="item-delete" onclick="deleteExamRecord('+r.id+')">Delete</button>'
            + '</td>'
            + '</tr>';
    }

    var html = '';
    if (active.length) {
        html += '<div style="padding:8px 12px;background:#f0fdf4;border-bottom:1px solid #e2e8f0;font-size:12px;font-weight:700;color:#15803d;">✅ ACTIVE EXAMS (appear in Import Marks dropdown)</div>';
        html += '<div class="tbl-wrap"><table style="width:100%;border-collapse:collapse;font-size:13px;">'
            + '<tr style="background:#eff6ff;"><th style="padding:8px 12px;text-align:left;font-size:11px;">Exam</th><th style="padding:8px 12px;font-size:11px;">Created</th><th style="padding:8px 12px;text-align:center;font-size:11px;">Marks</th><th style="padding:8px 12px;text-align:center;font-size:11px;">Status</th><th style="padding:8px 12px;font-size:11px;">Actions</th></tr>';
        active.forEach(function(r){ html += examRow(r); });
        html += '</table></div>';
    }
    if (closed.length) {
        html += '<div style="padding:8px 12px;background:#f8fafc;border-bottom:1px solid #e2e8f0;border-top:2px solid #e2e8f0;font-size:12px;font-weight:700;color:#6b7280;margin-top:4px;">🔒 CLOSED EXAMS</div>';
        html += '<div class="tbl-wrap"><table style="width:100%;border-collapse:collapse;font-size:13px;">'
            + '<tr style="background:#f8fafc;"><th style="padding:8px 12px;text-align:left;font-size:11px;">Exam</th><th style="padding:8px 12px;font-size:11px;">Created</th><th style="padding:8px 12px;text-align:center;font-size:11px;">Marks</th><th style="padding:8px 12px;text-align:center;font-size:11px;">Status</th><th style="padding:8px 12px;font-size:11px;">Actions</th></tr>';
        closed.forEach(function(r){ html += examRow(r); });
        html += '</table></div>';
    }
    container.innerHTML = html;
}

function toggleExamStatus(id) {
    var r = examRecords.find(function(r){ return r.id===id; });
    if (!r) return;
    r.status = r.status==='active' ? 'closed' : 'active';
    saveAllData();
    refreshAllDropdowns();
    renderExamRecordsList();
    showMessage('✓ Exam "'+r.name+'" is now '+r.status,'success');
}

function deleteExamRecord(id) {
    var r = examRecords.find(function(r){ return r.id===id; });
    if (!r) return;
    var hasMarks = exams[r.name] && exams[r.name].length > 0;
    var msg = hasMarks
        ? 'Delete exam "'+r.name+'" AND all its marks data? This cannot be undone.'
        : 'Delete exam "'+r.name+'"?';
    if (!confirm(msg)) return;
    // Remove marks data
    (exams[r.name]||[]).forEach(function(c){ delete examMarks[r.name+'_'+c]; });
    delete exams[r.name];
    // Remove record
    examRecords = examRecords.filter(function(rec){ return rec.id!==id; });
    saveAllData();
    refreshAllDropdowns();
    renderExamRecordsList();
    showMessage('✓ Exam deleted','success');
}

// Populate the import marks exam dropdown with active exams only
function refreshImportExamDropdown() {
    var activeExams = examRecords.filter(function(r){ return r.status==='active'; });
    ['importExamName','aioExamName'].forEach(function(id){
        var sel = document.getElementById(id); if(!sel) return;
        var cur = sel.value;
        sel.innerHTML = '<option value="">-- Select Active Exam --</option>';
        if (!activeExams.length) {
            var o = document.createElement('option');
            o.disabled = true; o.textContent = '(No active exams — create one in Exams tab)';
            sel.appendChild(o);
        } else {
            activeExams.forEach(function(r){
                var o = document.createElement('option');
                o.value = r.name;
                o.textContent = r.name + ' · ' + r.type + ' · ' + r.term + ' ' + r.year;
                if (r.name===cur) o.selected = true;
                sel.appendChild(o);
            });
        }
    });
    onImportExamChange();
    aioOnExamChange();
}

function onImportExamChange() {
    var val = document.getElementById('importExamName').value;
    var info = document.getElementById('importExamInfo');
    if (!val || !info) return;
    var r = examRecords.find(function(r){ return r.name===val; });
    if (r) {
        var hasMarks = exams[r.name] && exams[r.name].length;
        var markCount = 0;
        if (hasMarks) (exams[r.name]||[]).forEach(function(c){ markCount += Object.keys(examMarks[r.name+'_'+c]||{}).length; });
        info.textContent = r.type + ' · ' + r.term + ' ' + r.year + (markCount ? ' · ' + markCount + ' mark entries already saved' : ' · No marks yet');
        info.style.color = markCount ? '#b45309' : '#15803d';
    } else {
        info.textContent = '';
    }
}

// ═══════════════ MOST IMPROVED ═══════════════
function downloadMostImproved() {
    var exam1 = document.getElementById('improvedExam1').value;
    var exam2 = document.getElementById('improvedExam2').value;
    if (!exam1 || !exam2) { showMessage('❌ Select both exams','error'); return; }
    if (exam1 === exam2)  { showMessage('❌ Select two different exams','error'); return; }

    // Collect totalPct per learner per exam using meanPct
    function getExamTotals(exam) {
        var totals = {}; // {name: {stream, totalPct}}
        cleanCombos(exams[exam]||[]).forEach(function(combo){
            var p=parseCombo(combo), key=exam+'_'+combo;
            Object.entries(examMarks[key]||{}).forEach(function(e){
                var name=e[0], raw=e[1];
                if(!totals[name]) totals[name]={stream:p.stream,pcts:[]};
                totals[name].pcts.push(toPercent(raw,p.subject,p.stream));
            });
        });
        // Convert pcts array to mean
        Object.entries(totals).forEach(function(e){
            e[1].totalPct = Math.round(e[1].pcts.reduce(function(a,b){return a+b;},0)/e[1].pcts.length);
        });
        return totals;
    }

    var base  = getExamTotals(exam1); // earlier
    var later = getExamTotals(exam2); // more recent

    // Find learners present in both exams, compute deviation
    var deviations = []; // [{name, stream, grade, base, later, dev}]
    Object.entries(later).forEach(function(e){
        var name=e[0], d=e[1];
        if (base[name] === undefined) return; // not in both
        deviations.push({
            name: name,
            stream: d.stream,
            grade: gradeFromStream(d.stream),
            basePct: base[name].totalPct,
            laterPct: d.totalPct,
            dev: d.totalPct - base[name].totalPct
        });
    });

    if (!deviations.length) { showMessage('❌ No learners found in both exams','error'); return; }

    // Group by grade number, top 3 per grade by deviation
    var byGrade = {};
    deviations.forEach(function(d){
        if (!byGrade[d.grade]) byGrade[d.grade] = [];
        byGrade[d.grade].push(d);
    });

    var allSheets = [];
    var summaryRows = [['Grade','Rank','Learner','Stream',exam1+' Total %',exam2+' Total %','Improvement (+/-)']];

    Object.keys(byGrade).sort().forEach(function(g){
        var group = byGrade[g].sort(function(a,b){ return b.dev - a.dev; });
        var top3  = group.slice(0,3);

        // Summary sheet rows
        top3.forEach(function(d, idx){
            summaryRows.push([
                'Grade '+g,
                idx+1,
                d.name,
                d.stream,
                d.basePct+'%',
                d.laterPct+'%',
                (d.dev>=0?'+':'')+d.dev+'%'
            ]);
        });
        summaryRows.push([]); // spacer

        // Per-grade sheet
        var sheetRows = [
            ['Most Improved Learners — Grade '+g],
            ['Based on: '+exam1+' → '+exam2],
            [],
            ['Rank','Learner','Stream',exam1+' Total %',exam2+' Total %','Improvement','Level Change']
        ];
        top3.forEach(function(d, idx){
            var g1=getGrade(d.basePct)||'—', g2=getGrade(d.laterPct)||'—';
            sheetRows.push([
                idx+1,
                d.name,
                d.stream,
                d.basePct+'%',
                d.laterPct+'%',
                (d.dev>=0?'+':'')+d.dev+'%',
                g1+' → '+g2
            ]);
        });
        sheetRows.push([]);
        sheetRows.push(['All learners in Grade '+g+' ('+group.length+' total), ranked by improvement:']);
        sheetRows.push(['Rank','Learner','Stream',exam1+' Total %',exam2+' Total %','Improvement']);
        group.forEach(function(d,idx){
            sheetRows.push([idx+1,d.name,d.stream,d.basePct+'%',d.laterPct+'%',(d.dev>=0?'+':'')+d.dev+'%']);
        });

        allSheets.push({name:'Grade '+g, data:sheetRows});
    });

    // Summary sheet first
    allSheets.unshift({name:'Summary', data:summaryRows});

    var filename = 'MostImproved_'+exam1.replace(/\s/g,'_')+'_to_'+exam2.replace(/\s/g,'_')+'.xlsx';
    exportXLSX(allSheets, filename, exam1+' → '+exam2);
    showMessage('✓ Most Improved downloaded — top 3 per grade across all streams','success');
}

// ═══════════════ TAB 14: ALL-IN-ONE IMPORT ═══════════════
var aioWB = null; // cached workbook

function aioOpenTab() {
    aioOnExamChange();
    // Reset if no workbook loaded
    if (!aioWB) {
        document.getElementById('aioSheetMapper').style.display = 'none';
        document.getElementById('aioFile').value = '';
    }
}

function aioOnExamChange() {
    var val = document.getElementById('aioExamName').value;
    var info = document.getElementById('aioExamInfo');
    if (!info) return;
    if (!val) { info.textContent = ''; return; }
    var r = examRecords.find(function(r){ return r.name===val; });
    if (r) {
        var markCount = 0;
        (exams[r.name]||[]).forEach(function(c){ markCount += Object.keys(examMarks[r.name+'_'+c]||{}).length; });
        info.textContent = r.type+' · '+r.term+' '+r.year+(markCount?' · '+markCount+' mark entries already saved':' · No marks yet');
        info.style.color = markCount ? '#b45309' : '#15803d';
    } else { info.textContent = ''; }
}

function aioLoadWorkbook() {
    var file = document.getElementById('aioFile').files[0];
    document.getElementById('aioSheetMapper').style.display = 'none';
    document.getElementById('aioPreviewArea').innerHTML = '';
    document.getElementById('aioImportStatus').textContent = '';
    aioWB = null;
    if (!file) return;

    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            aioWB = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            aioRenderMapper();
        } catch(err) { showMessage('❌ Error reading file: '+err.message, 'error'); }
    };
    reader.readAsArrayBuffer(file);
}

function aioRenderMapper() {
    if (!aioWB) return;
    var sheetNames = aioWB.SheetNames;
    document.getElementById('aioSheetMapper').style.display = 'block';

    var html = '<div class="tbl-wrap"><table style="font-size:13px;">'
        + '<tr style="background:#eff6ff;">'
        + '<th style="padding:9px 12px;">Sheet Name (in workbook)</th>'
        + '<th style="padding:9px 12px;">→ Assign to Stream</th>'
        + '<th style="padding:9px 12px;text-align:center;">Preview</th>'
        + '</tr>';

    sheetNames.forEach(function(sheet, idx) {
        // Try auto-match: if sheet name matches an existing class exactly or fuzzy
        var autoMatch = '';
        classes.forEach(function(c) {
            if (c.toUpperCase() === sheet.trim().toUpperCase()) autoMatch = c;
        });
        if (!autoMatch) {
            // fuzzy: sheet "9A" matches class "9A", "Grade 9A" matches "9A" etc
            classes.forEach(function(c) {
                if (sheet.toUpperCase().indexOf(c.toUpperCase()) >= 0 ||
                    c.toUpperCase().indexOf(sheet.trim().toUpperCase()) >= 0) {
                    if (!autoMatch) autoMatch = c;
                }
            });
        }

        var opts = '<option value="">-- Skip this sheet --</option>';
        classes.slice().sort().forEach(function(c) {
            opts += '<option value="'+c+'"'+(c===autoMatch?' selected':'')+'>'+c+'</option>';
        });
        opts += '<option value="__new__">➕ Create new stream named "'+sheet.trim()+'"</option>';

        html += '<tr>'
            + '<td style="padding:9px 12px;font-weight:600;">📄 '+sheet+'</td>'
            + '<td style="padding:9px 12px;">'
            +   '<select id="aioMap_'+idx+'" onchange="aioPreviewSheet('+idx+')" style="padding:7px 10px;border:1px solid #d1d5db;border-radius:4px;font-size:13px;width:180px;">'
            +   opts+'</select>'
            + '</td>'
            + '<td style="padding:9px 12px;text-align:center;">'
            +   '<button class="btn" style="padding:5px 12px;font-size:11px;margin:0;" onclick="aioPreviewSheet('+idx+')">👁 Preview</button>'
            + '</td>'
            + '</tr>';
    });

    html += '</table></div>';
    document.getElementById('aioMappingRows').innerHTML = html;

    // Auto-preview all mapped sheets
    sheetNames.forEach(function(sheet, idx) {
        aioPreviewSheet(idx);
    });
}

function aioPreviewSheet(idx) {
    if (!aioWB) return;
    var sheetName = aioWB.SheetNames[idx];
    var streamSel = document.getElementById('aioMap_'+idx);
    var stream = streamSel ? streamSel.value : '';
    if (!stream || stream === '__new__') {
        stream = stream === '__new__' ? sheetName.trim() : '';
    }

    var rows = XLSX.utils.sheet_to_json(aioWB.Sheets[sheetName], { header: 1 });
    if (rows.length < 5) return;

    var colMap = parseMarkHeaders(rows); // subjects from row 5 (index 4), col E+ (index 4+)
    var subjHeaders = colMap.map(function(c){ return c.logicalName; });
    var uniqueSubjs = [...new Set(subjHeaders)];

    // Build per-sheet preview
    var previewId = 'aioPreview_'+idx;
    var existing = document.getElementById(previewId);
    if (!existing) {
        var div = document.createElement('div');
        div.id = previewId;
        document.getElementById('aioPreviewArea').appendChild(div);
        existing = div;
    }

    var learnerCount = 0;
    for (var i = 5; i < rows.length; i++) {
        if (rows[i] && rows[i][1]) learnerCount++;
    }

    var html = '<div class="block" style="margin-bottom:12px;">'
        + '<div class="block-hdr" style="font-size:12px;">'
        +   '📄 <strong>'+sheetName+'</strong>'
        +   (stream ? ' → <span class="stream-tag">'+stream+'</span>' : ' <span style="color:#dc2626;">⚠ Not mapped</span>')
        +   ' &nbsp;·&nbsp; '+learnerCount+' learners'
        +   (uniqueSubjs.length ? ' &nbsp;·&nbsp; Subjects: <strong>'+uniqueSubjs.join(', ')+'</strong>' : ' <span style="color:#dc2626;">⚠ No subject headers found in Row 5</span>')
        + '</div>';

    if (learnerCount > 0 && colMap.length > 0) {
        html += '<div class="tbl-wrap"><table style="font-size:12px;">'
            + '<tr style="background:#eff6ff;"><th style="padding:6px 10px;">#</th><th style="padding:6px 10px;">Name (Col B)</th><th style="padding:6px 10px;">Gender (Col C)</th>';
        uniqueSubjs.forEach(function(s){ html += '<th style="padding:6px 10px;">'+subjLabel(s)+'</th>'; });
        html += '</tr>';

        var shown = 0;
        for (var i = 5; i < rows.length && shown < 5; i++) {
            var r = rows[i]; if (!r || !r[1]) continue; shown++;
            var rowMarks = computeRowMarks(r, colMap);
            html += '<tr>'
                + '<td style="padding:6px 10px;color:#6b7280;">'+(r[0]||i-4)+'</td>'
                + '<td style="padding:6px 10px;font-weight:600;">'+String(r[1]||'')+'</td>'
                + '<td style="padding:6px 10px;">'+String(r[2]||'')+'</td>';
            uniqueSubjs.forEach(function(s){
                var val = rowMarks[s];
                var pct = val !== undefined ? Math.round(toPercent(val,s,stream))+'%' : '—';
                html += '<td style="padding:6px 10px;color:#1e40af;font-weight:600;">'+pct+'</td>';
            });
            html += '</tr>';
        }
        if (learnerCount > 5) html += '<tr><td colspan="'+(3+uniqueSubjs.length)+'" style="padding:6px 10px;color:#6b7280;font-size:11px;">... and '+(learnerCount-5)+' more</td></tr>';
        html += '</table></div>';
    }
    html += '</div>';
    existing.innerHTML = html;
}

function aioImportAll() {
    var examName = document.getElementById('aioExamName').value;
    if (!examName) { showMessage('❌ Select an active exam first', 'error'); return; }
    if (!aioWB)    { showMessage('❌ Upload a workbook first', 'error'); return; }

    var statusEl = document.getElementById('aioImportStatus');
    statusEl.textContent = 'Importing…';

    var sheetNames = aioWB.SheetNames;
    var totalLearners = 0, totalMarks = 0, streamsImported = [];
    var warnings = [];

    if (!exams[examName]) exams[examName] = [];

    sheetNames.forEach(function(sheetName, idx) {
        var sel = document.getElementById('aioMap_'+idx);
        var stream = sel ? sel.value : '';
        if (!stream) return; // skip unmapped sheets

        // Create new stream if "__new__" selected
        if (stream === '__new__') stream = sheetName.trim();

        var rows = XLSX.utils.sheet_to_json(aioWB.Sheets[sheetName], { header: 1 });
        if (rows.length < 6) { warnings.push('Sheet "'+sheetName+'": fewer than 6 rows, skipped.'); return; }

        var colMap = parseMarkHeaders(rows);
        if (!colMap.length) { warnings.push('Sheet "'+sheetName+'": no subject headers in Row 5 from Col E, skipped marks.'); }

        // Ensure stream exists in classes
        if (!classes.includes(stream)) classes.push(stream);
        if (!learners[stream]) learners[stream] = [];

        var sheetLearners = 0, sheetMarks = 0;

        for (var i = 5; i < rows.length; i++) {
            var r = rows[i];
            if (!r || !r[1]) continue;

            var name   = String(r[1]||'').trim();
            var gender = String(r[2]||'').trim();
            if (!name) continue;

            // Import learner
            if (!learners[stream].find(function(l){ return l.name===name; })) {
                learners[stream].push({ name: name, gender: gender });
                sheetLearners++;
            }

            // Import marks
            if (colMap.length) {
                var rowMarks = computeRowMarks(r, colMap);
                Object.entries(rowMarks).forEach(function(e) {
                    var subj = e[0], raw = e[1];
                    var key = examName+'_'+stream+'_'+subj;
                    if (!examMarks[key]) examMarks[key] = {};
                    examMarks[key][name] = raw;
                    sheetMarks++;
                    var combo = stream+'_'+subj;
                    if (!exams[examName].includes(combo)) exams[examName].push(combo);
                });
            }
        }

        totalLearners += sheetLearners;
        totalMarks    += sheetMarks;
        streamsImported.push(stream);
    });

    if (!streamsImported.length) {
        statusEl.textContent = '';
        showMessage('❌ No sheets were mapped to streams. Assign each sheet to a stream first.', 'error');
        return;
    }

    saveAllData();
    refreshAllDropdowns();

    var msg = '✓ Imported '+totalLearners+' new learners and '+totalMarks+' marks across '
        + streamsImported.length+' stream(s): '+streamsImported.join(', ');
    statusEl.textContent = msg;
    statusEl.style.color = '#15803d';

    if (warnings.length) {
        warnings.forEach(function(w){ showMessage('⚠️ '+w, 'info'); });
    } else {
        showMessage(msg, 'success');
    }

    // Reset file input, keep exam selected and mapping visible for re-use
    document.getElementById('aioFile').value = '';
    aioWB = null;
    document.getElementById('aioSheetMapper').style.display = 'none';
    document.getElementById('aioPreviewArea').innerHTML = '';

    // Auto-navigate to analysis
    setTimeout(function(){ switchTab(5, null); }, 900);
}

// ═══════════════ TOP PERFORMERS ═══════════════
function downloadTopPerformers() {
    var exam = getDownloadExam(); if (!exam) return;
    var combos = cleanCombos(exams[exam]||[]);
    if (!combos.length) { showMessage('❌ No data for this exam','error'); return; }

    var allSubjects = orderedSubjects(combos);

    // Build per-learner data: {name:{stream, grade, pct:{subj:pct}}}
    var learnerData = {};
    combos.forEach(function(combo){
        var p = parseCombo(combo), key = exam+'_'+combo;
        Object.entries(examMarks[key]||{}).forEach(function(e){
            var name=e[0], raw=e[1];
            if (!learnerData[name]) learnerData[name] = {stream:p.stream, grade:gradeFromStream(p.stream), pct:{}};
            learnerData[name].pct[p.subject] = toPercent(raw, p.subject, p.stream);
        });
    });

    // Group learners by grade
    var byGrade = {};
    Object.entries(learnerData).forEach(function(e){
        var name=e[0], d=e[1];
        if (!byGrade[d.grade]) byGrade[d.grade] = [];
        byGrade[d.grade].push({name:name, stream:d.stream, pct:d.pct});
    });

    var sheets = [];
    // Summary sheet — one row per grade per category (subject + overall)
    var summaryRows = [
        ['Top 3 Performers per Grade — Exam: '+exam],
        [],
        ['Grade','Category','Rank','Learner','Stream','Score %','Level']
    ];

    Object.keys(byGrade).sort().forEach(function(g){
        var group = byGrade[g];

        // Per-grade detail sheet rows
        var detailRows = [
            ['Grade '+g+' — Top 3 Performers'],
            ['Exam: '+exam],
            ['Streams: '+[...new Set(group.map(function(d){return d.stream;}))].sort().join(', ')],
            ['Total learners in grade: '+group.length],
            []
        ];

        // ── Overall top 3 ──
        var overallRanked = group
            .map(function(d){ return {name:d.name, stream:d.stream, score:parseFloat(meanPct(d.pct).toFixed(2))}; })
            .sort(function(a,b){ return b.score-a.score; });

        detailRows.push(['🏆 OVERALL — Top 3 (by Total %)']);
        detailRows.push(['Rank','Learner','Stream','Total %','Level']);
        overallRanked.slice(0,3).forEach(function(d,i){
            var lv = getGrade(d.score)||'—';
            detailRows.push([i+1, d.name, d.stream, d.score+'%', lv]);
            summaryRows.push(['Grade '+g, 'Overall', i+1, d.name, d.stream, d.score+'%', lv]);
        });
        detailRows.push([]);
        summaryRows.push([]);

        // ── Per-subject top 3 ──
        // Only use subjects present in this grade
        var gradeSubjects = orderedSubjects(
            combos.filter(function(c){ return gradeFromStream(parseCombo(c).stream)===g; })
        );

        gradeSubjects.forEach(function(subj){
            // Only include learners who actually have a mark for this subject
            var subjRanked = group
                .filter(function(d){ return d.pct[subj] !== undefined; })
                .map(function(d){ return {name:d.name, stream:d.stream, score:Math.round(d.pct[subj])}; })
                .sort(function(a,b){ return b.score-a.score; });

            if (!subjRanked.length) return;

            var label = subjLabel(subj)+(isPaperSubject(subj)?' (P1+P2)':'');
            detailRows.push(['📖 '+label+' — Top 3']);
            detailRows.push(['Rank','Learner','Stream',label+' %','Level']);
            subjRanked.slice(0,3).forEach(function(d,i){
                var lv = getGrade(d.score)||'—';
                detailRows.push([i+1, d.name, d.stream, d.score+'%', lv]);
                summaryRows.push(['Grade '+g, label, i+1, d.name, d.stream, d.score+'%', lv]);
            });
            detailRows.push([]);
            summaryRows.push([]);
        });

        sheets.push({name:'Grade '+g, data:detailRows});
    });

    // Summary sheet first
    sheets.unshift({name:'Summary', data:summaryRows});

    exportXLSX(sheets, exam+'_TopPerformers.xlsx', exam);
    showMessage('✓ Top Performers downloaded — top 3 per subject and overall per grade','success');
}

// ═══════════════ PERSISTENCE ═══════════════
function saveAllData() {
    function s(k,v){ localStorage.setItem(k,JSON.stringify(v)); }
    s('schoolSetup',schoolSetup); s('classes',classes); s('subjects',subjects);
    s('grades',grades); s('exams',exams); s('examMarks',examMarks); s('learners',learners);
    s('paperConfig',paperConfig); s('teachers',teachers); s('examRecords',examRecords);
    s('gradeSubjectConfig',gradeSubjectConfig);
}
function loadAllData() {
    function p(k,fb){ try{ return JSON.parse(localStorage.getItem(k))||fb; }catch(e){ return fb; } }
    schoolSetup=p('schoolSetup',{name:'',motto:'',address:'',logo:''});
    classes=p('classes',[]); subjects=p('subjects',[]); grades=p('grades',[]);
    exams=p('exams',{}); examMarks=p('examMarks',{}); learners=p('learners',{});
    paperConfig=p('paperConfig',{eng:{p1Max:50,p2Max:50},kis:{p1Max:50,p2Max:50}});
    teachers=p('teachers',{principal:{name:'',title:'Principal',sig:''},streams:{}});
    examRecords=p('examRecords',[]);
    gradeSubjectConfig=p('gradeSubjectConfig',{default:{subjects:{},eng:100,kis:100}});
}

// ═══════════════════════════════════════════════════════
//  AUTH SYSTEM
//  Users stored in localStorage as: csUsers = [{username, hash, role}]
//  Roles: 'admin' | 'teacher'
//  Session stored in sessionStorage so logout on tab close
// ═══════════════════════════════════════════════════════

// Simple hash — SHA-like using string manipulation (no crypto dependency)
function simpleHash(str) {
    var hash = 5381;
    for (var i = 0; i < str.length; i++) {
        hash = ((hash << 5) + hash) ^ str.charCodeAt(i);
        hash = hash & 0x7fffffff; // keep positive 31-bit
    }
    return 'cs_' + hash.toString(36) + '_' + str.length;
}

function getUsers() {
    try { return JSON.parse(localStorage.getItem('csUsers')) || []; } catch(e) { return []; }
}
function saveUsers(users) {
    localStorage.setItem('csUsers', JSON.stringify(users));
}

// Seed default admin if no users exist
function ensureDefaultAdmin() {
    var users = getUsers();
    if (!users.length) {
        users.push({ username: 'admin', hash: simpleHash('admin123'), role: 'admin' });
        saveUsers(users);
    }
}

function doLogin() {
    var username = document.getElementById('loginUsername').value.trim();
    var password = document.getElementById('loginPassword').value;
    var errEl    = document.getElementById('loginError');

    if (!username || !password) {
        errEl.textContent = 'Please enter both username and password.';
        errEl.classList.add('show'); return;
    }

    var users = getUsers();
    var user  = users.find(function(u){ return u.username.toLowerCase() === username.toLowerCase(); });

    if (!user || user.hash !== simpleHash(password)) {
        errEl.textContent = 'Incorrect username or password.';
        errEl.classList.add('show');
        document.getElementById('loginPassword').value = '';
        return;
    }

    // Successful login
    errEl.classList.remove('show');
    sessionStorage.setItem('csSession', JSON.stringify({ username: user.username, role: user.role }));
    showApp(user);
}

function doLogout() {
    if (!confirm('Sign out of CoreSpark?')) return;
    sessionStorage.removeItem('csSession');
    document.getElementById('appRoot').classList.remove('visible');
    document.getElementById('loginScreen').classList.remove('hidden');
    document.getElementById('loginUsername').value = '';
    document.getElementById('loginPassword').value = '';
    document.getElementById('loginError').classList.remove('show');
    document.getElementById('loggedInUser').textContent = '';
}

// ── Role helpers ──
function getRole() {
    try { return JSON.parse(sessionStorage.getItem('csSession')||'{}').role || 'teacher'; } catch(e){ return 'teacher'; }
}
function getUsername() {
    try { return JSON.parse(sessionStorage.getItem('csSession')||'{}').username || ''; } catch(e){ return ''; }
}
function isAdmin() { return getRole()==='admin'; }

// ── Tab visibility by role ──
// Admin: all tabs | Teacher: Enter Marks + Past Results (read-only)
var ADMIN_TABS = [
    {i:0,  label:'🏫 School'},
    {i:12, label:'👨‍🏫 Teachers'},
    {i:2,  label:'📖 Subjects'},
    {i:3,  label:'⚙️ Levels'},
    {i:13, label:'📝 Exams'},
    {i:14, label:'⚡ Import All-in-One'},
    {i:8,  label:'👥 Import Learners'},
    {i:10, label:'🗂️ Learners Store'},
    {i:9,  label:'📊 Import Marks'},
    {i:11, label:'📑 Marks Store'},
    {i:5,  label:'📈 Analysis'},
    {i:6,  label:'📋 Past Results'},
    {i:7,  label:'📥 Downloads'},
    {i:1,  label:'📚 Classes'},
    {i:4,  label:'📤 Enter Marks'}
];
var TEACHER_TABS = [
    {i:4,  label:'📤 Enter Marks'},
    {i:6,  label:'📋 Past Results'}
];

function buildTabBar(role) {
    var tabs = role==='admin' ? ADMIN_TABS : TEACHER_TABS;
    var bar  = document.querySelector('.tabs');
    bar.innerHTML = '';
    tabs.forEach(function(t, idx){
        var btn = document.createElement('button');
        btn.className = 'tab' + (idx===0?' active':'');
        btn.textContent = t.label;
        btn.setAttribute('onclick','switchTab('+t.i+',this)');
        bar.appendChild(btn);
    });
    // Activate first tab
    switchTab(tabs[0].i, bar.querySelector('.tab.active'));
}

function showApp(user) {
    document.getElementById('loginScreen').classList.add('hidden');
    document.getElementById('appRoot').classList.add('visible');
    document.getElementById('loggedInUser').textContent = '👤 '+user.username+' ('+user.role+')';

    // Build role-appropriate tab bar
    buildTabBar(user.role);

    // Initialise app data
    loadAllData();
    updateSchoolDisplay();
    refreshExamSelects();
    loadPaperConfigUI();
    refreshImportExamDropdown();

    // Teacher banner
    var banner = document.getElementById('teacherBanner');
    if (banner) {
        banner.style.display = user.role==='teacher' ? 'block' : 'none';
    }

    showMessage('✓ Welcome, '+user.username+'! CoreSpark Ready.', 'success');
    renderUserManager();
    applyRoleUI();
    if (user.role==='teacher') {
        var enterBtn=document.querySelector('.tab[onclick="switchTab(4,this)"]');
        if(enterBtn) switchTab(4,enterBtn);
    }
}

function checkSession() {
    ensureDefaultAdmin();
    var session = sessionStorage.getItem('csSession');
    if (session) {
        try {
            var user = JSON.parse(session);
            // Verify user still exists
            var users = getUsers();
            var found = users.find(function(u){ return u.username === user.username; });
            if (found) { showApp(found); return; }
        } catch(e) {}
    }
    // Show login screen
    document.getElementById('loginScreen').classList.remove('hidden');
    document.getElementById('loginUsername').focus();
}

// ═══════════════════════════════════════════════════════
//  USER MANAGEMENT (accessible from School/Admin tab)
// ═══════════════════════════════════════════════════════
function renderUserManager() {
    var el = document.getElementById('userManagerContent');
    if (!el) return;
    var session = JSON.parse(sessionStorage.getItem('csSession')||'{}');
    if (session.role !== 'admin') {
        el.innerHTML = '<p style="color:#6b7280;font-size:13px;">Only admin users can manage accounts.</p>'; return;
    }
    var users = getUsers();
    var html = '<div class="tbl-wrap"><table style="font-size:13px;">'
        + '<tr style="background:#eff6ff;"><th style="padding:9px 12px;">Username</th><th style="padding:9px 12px;">Role</th><th style="padding:9px 12px;">Actions</th></tr>';
    users.forEach(function(u, idx){
        html += '<tr>'
            + '<td style="padding:9px 12px;font-weight:600;">'+u.username+'</td>'
            + '<td style="padding:9px 12px;"><span style="background:'+(u.role==='admin'?'#dbeafe':'#dcfce7')+';color:'+(u.role==='admin'?'#1e3a8a':'#14532d')+';padding:2px 8px;border-radius:8px;font-size:11px;font-weight:700;">'+u.role+'</span></td>'
            + '<td style="padding:9px 12px;">'
            + (u.username!=='admin'?'<button class="item-delete" onclick="deleteUser(\''+u.username+'\')">Delete</button> ':' ')
            + '<button class="btn" style="font-size:11px;padding:4px 10px;margin:0;" onclick="resetPassword(\''+u.username+'\')">Reset Password</button>'
            + '</td></tr>';
    });
    html += '</table></div>';
    el.innerHTML = html;
}

function addUser() {
    var uname = document.getElementById('newUsername').value.trim();
    var pwd   = document.getElementById('newPassword').value;
    var role  = document.getElementById('newUserRole').value;
    if (!uname || !pwd) { showMessage('❌ Enter username and password','error'); return; }
    if (pwd.length < 4) { showMessage('❌ Password must be at least 4 characters','error'); return; }
    var users = getUsers();
    if (users.find(function(u){ return u.username.toLowerCase()===uname.toLowerCase(); })){
        showMessage('❌ Username already exists','error'); return;
    }
    users.push({ username: uname, hash: simpleHash(pwd), role: role });
    saveUsers(users);
    document.getElementById('newUsername').value = '';
    document.getElementById('newPassword').value = '';
    renderUserManager();
    showMessage('✓ User "'+uname+'" created','success');
}

function deleteUser(username) {
    if (!confirm('Delete user "'+username+'"?')) return;
    var users = getUsers().filter(function(u){ return u.username !== username; });
    saveUsers(users);
    renderUserManager();
    showMessage('✓ User deleted','success');
}

function resetPassword(username) {
    var newPwd = prompt('New password for "'+username+'" (min 4 characters):');
    if (!newPwd) return;
    if (newPwd.length < 4) { showMessage('❌ Password too short','error'); return; }
    var users = getUsers();
    var u = users.find(function(u){ return u.username===username; });
    if (u) { u.hash = simpleHash(newPwd); saveUsers(users); showMessage('✓ Password reset for "'+username+'"','success'); }
}

// ═══════════════════════════════════════════════════════
//  INIT — check session instead of loading app directly
// ═══════════════════════════════════════════════════════
window.onload = function() { checkSession(); };
</script>
</div><!-- /appRoot -->
</body>
</html>
