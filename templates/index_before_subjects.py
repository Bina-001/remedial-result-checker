<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>Remedial Result Checker</title>

    <style>

        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            font-family: Arial, sans-serif;
            background: #f4f7fb;
            color: #172033;
        }

        .container {
            width: 95%;
            max-width: 1200px;
            margin: 25px auto;
        }

        .header {
            background: #172033;
            color: white;
            padding: 22px;
            border-radius: 14px;
            margin-bottom: 20px;
        }

        .header h1 {
            margin: 0 0 8px;
        }

        .header p {
            margin: 0;
            opacity: 0.8;
        }

        .card {
            background: white;
            padding: 20px;
            border-radius: 14px;
            margin-bottom: 20px;
            box-shadow: 0 3px 12px rgba(0,0,0,0.07);
        }

        .upload-row {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            align-items: center;
        }

        input[type="file"] {
            flex: 1;
            min-width: 220px;
        }

        button,
        .download-btn {
            border: none;
            border-radius: 8px;
            padding: 11px 16px;
            cursor: pointer;
            font-weight: bold;
            text-decoration: none;
            display: inline-block;
        }

        .primary {
            background: #2563eb;
            color: white;
        }

        .success {
            background: #16a34a;
            color: white;
        }

        .warning {
            background: #f59e0b;
            color: white;
        }

        .dark {
            background: #172033;
            color: white;
        }

        button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .message {
            margin-top: 15px;
            padding: 12px;
            border-radius: 8px;
            background: #eef2ff;
        }

        .summary-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 10px;
        }

        .summary-box {
            background: #f8fafc;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
        }

        .summary-number {
            font-size: 25px;
            font-weight: bold;
            margin-top: 5px;
        }

        .progress-area {
            margin-top: 18px;
        }

        .progress-track {
            width: 100%;
            height: 20px;
            background: #e5e7eb;
            border-radius: 20px;
            overflow: hidden;
        }

        .progress-bar {
            height: 100%;
            width: 0%;
            background: #2563eb;
            transition: width 0.3s;
        }

        .progress-info {
            display: flex;
            justify-content: space-between;
            margin-top: 8px;
            font-size: 14px;
        }

        .actions {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .table-wrapper {
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 850px;
        }

        th,
        td {
            padding: 10px;
            border-bottom: 1px solid #e5e7eb;
            text-align: left;
            font-size: 14px;
        }

        th {
            background: #f8fafc;
        }

        .status {
            font-weight: bold;
        }

        .status-pass {
            color: #15803d;
        }

        .status-fail {
            color: #dc2626;
        }

        .status-disqualified {
            color: #b45309;
        }

        .status-error {
            color: #dc2626;
        }

        .status-miss {
            color: #6b7280;
        }

        .status-other {
            color: #2563eb;
        }

        @media (max-width: 900px) {

            .summary-grid {
                grid-template-columns: repeat(4, 1fr);
            }

        }

        @media (max-width: 600px) {

            .container {
                width: 92%;
                margin: 15px auto;
            }

            .summary-grid {
                grid-template-columns: repeat(2, 1fr);
            }

            .header h1 {
                font-size: 23px;
            }

            button,
            .download-btn {
                width: 100%;
            }

            .actions {
                display: grid;
                grid-template-columns: 1fr;
            }

        }

    </style>

</head>


<body>

<div class="container">


    <div class="header">

        <h1>
            Remedial Result Checker
        </h1>

        <p>
            Ministry Result Checking System
        </p>

    </div>


    <div class="card">

        <h2>
            1. Upload Student Roster
        </h2>

        <form id="uploadForm">

            <div class="upload-row">

                <input
                    type="file"
                    id="fileInput"
                    accept=".xlsx,.csv"
                    required
                >

                <button
                    type="submit"
                    class="primary"
                    id="uploadButton"
                >
                    Upload Roster
                </button>

            </div>

        </form>

        <div
            id="uploadMessage"
            class="message"
        >
            Please upload a CSV or XLSX roster.
        </div>

    </div>


    <div class="card">

        <h2>
            2. Summary
        </h2>

        <div class="summary-grid">

            <div class="summary-box">
                <div>Total</div>
                <div
                    class="summary-number"
                    id="totalCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>Completed</div>
                <div
                    class="summary-number"
                    id="completedCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>Remaining</div>
                <div
                    class="summary-number"
                    id="remainingCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>PASS</div>
                <div
                    class="summary-number"
                    id="passCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>FAIL</div>
                <div
                    class="summary-number"
                    id="failCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>Disqualified</div>
                <div
                    class="summary-number"
                    id="disqualifiedCount"
                >
                    0
                </div>
            </div>

            <div class="summary-box">
                <div>Error</div>
                <div
                    class="summary-number"
                    id="errorCount"
                >
                    0
                </div>
            </div>

        </div>


        <div class="progress-area">

            <div class="progress-track">

                <div
                    id="progressBar"
                    class="progress-bar"
                ></div>

            </div>

            <div class="progress-info">

                <span id="progressText">
                    0%
                </span>

                <span id="etaText">
                    Waiting...
                </span>

            </div>

        </div>

    </div>


    <div class="card">

        <h2>
            3. Actions
        </h2>

        <div class="actions">

            <button
                id="testButton"
                class="warning"
                disabled
            >
                Test One Student
            </button>

            <button
                id="checkAllButton"
                class="success"
                disabled
            >
                Check All Students
            </button>

            <button
                id="resumeButton"
                class="dark"
                disabled
            >
                Resume Checking
            </button>

            <a
                href="/export/csv"
                class="download-btn primary"
            >
                Export CSV
            </a>

            <a
                href="/export/xlsx"
                class="download-btn dark"
            >
                Export Excel
            </a>

        </div>

        <div
            id="actionMessage"
            class="message"
        >
            Upload a roster to begin.
        </div>

    </div>


    <div class="card">

        <h2>
            4. Results
        </h2>

        <div class="table-wrapper">

            <table>

                <thead>

                    <tr>

                        <th>No.</th>
                        <th>Username</th>
                        <th>Name</th>
                        <th>Institution</th>
                        <th>Stream</th>
                        <th>Status</th>
                        <th>Score</th>
                        <th>Message</th>

                    </tr>

                </thead>

                <tbody id="resultsBody"></tbody>

            </table>

        </div>

    </div>


</div>


<script>

let students = [];

let checkedResults = [];

let checking = false;


function escapeHtml(value) {

    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

}


function sleep(ms) {

    return new Promise(
        resolve => setTimeout(resolve, ms)
    );

}


function getStatusClass(status) {

    const value =
        String(status || "").toLowerCase();

    if (value === "pass") {
        return "status-pass";
    }

    if (value === "fail") {
        return "status-fail";
    }

    if (value === "disqualified") {
        return "status-disqualified";
    }

    if (value === "error") {
        return "status-error";
    }

    if (value === "miss") {
        return "status-miss";
    }

    return "status-other";

}


function renderResults() {

    const body =
        document.getElementById(
            "resultsBody"
        );

    body.innerHTML = "";


    checkedResults.forEach(
        (result, index) => {

            const row =
                document.createElement("tr");


            row.innerHTML = `

                <td>
                    ${index + 1}
                </td>

                <td>
                    ${escapeHtml(
                        result.username || ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        result.full_name || ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        result.institution || ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        result.stream || ""
                    )}
                </td>

                <td
                    class="status ${getStatusClass(
                        result.status
                    )}"
                >
                    ${escapeHtml(
                        result.status || ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        result.score || ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        result.message || ""
                    )}
                </td>

            `;


            body.appendChild(row);

        }
    );

}


async function updateSummary() {

    try {

        const response =
            await fetch("/summary");

        const data =
            await response.json();


        if (!data.success) {
            return;
        }


        document.getElementById(
            "totalCount"
        ).textContent =
            data.total;


        document.getElementById(
            "completedCount"
        ).textContent =
            data.completed;


        document.getElementById(
            "remainingCount"
        ).textContent =
            data.remaining;


        document.getElementById(
            "passCount"
        ).textContent =
            data.pass;


        document.getElementById(
            "failCount"
        ).textContent =
            data.fail;


        document.getElementById(
            "disqualifiedCount"
        ).textContent =
            data.disqualified;


        document.getElementById(
            "errorCount"
        ).textContent =
            data.error;


        const total =
            Number(data.total || 0);


        const completed =
            Number(data.completed || 0);


        const percent =
            total > 0
                ? Math.round(
                    completed / total * 100
                )
                : 0;


        document.getElementById(
            "progressBar"
        ).style.width =
            percent + "%";


        document.getElementById(
            "progressText"
        ).textContent =
            percent + "%";


    } catch (error) {

        console.log(error);

    }

}


async function loadStudents() {

    try {

        const response =
            await fetch("/students");


        const data =
            await response.json();


        if (data.success) {

            students =
                data.students || [];

        }

    } catch (error) {

        console.log(error);

    }

}


async function loadSavedResults() {

    try {

        const response =
            await fetch("/results");


        const data =
            await response.json();


        if (data.success) {

            checkedResults =
                data.results || [];

            renderResults();

        }

    } catch (error) {

        console.log(error);

    }

}


document
    .getElementById("uploadForm")
    .addEventListener(
        "submit",
        async function(event) {

            event.preventDefault();


            const fileInput =
                document.getElementById(
                    "fileInput"
                );


            const message =
                document.getElementById(
                    "uploadMessage"
                );


            const file =
                fileInput.files[0];


            if (!file) {

                message.textContent =
                    "Please select a file.";

                return;

            }


            const formData =
                new FormData();


            formData.append(
                "file",
                file
            );


            message.textContent =
                "Uploading roster...";


            try {

                const response =
                    await fetch(
                        "/upload",
                        {
                            method: "POST",
                            body: formData
                        }
                    );


                const data =
                    await response.json();


                if (!data.success) {

                    message.textContent =
                        data.message ||
                        "Upload failed.";

                    return;

                }


                students = [];

                checkedResults = [];


                message.innerHTML = `

                    <b>
                        ${escapeHtml(
                            data.message
                        )}
                    </b>

                    <br>

                    Total students:
                    ${data.total_students}

                    <br>

                    Ready students:
                    ${data.ready_students}

                    <br>

                    Missing usernames:
                    ${data.missing_usernames}

                    <br>

                    Test username:
                    ${escapeHtml(
                        data.test_username
                    )}

                `;


                await loadStudents();

                await updateSummary();

                renderResults();


                document.getElementById(
                    "testButton"
                ).disabled = false;


                document.getElementById(
                    "checkAllButton"
                ).disabled = false;


                document.getElementById(
                    "resumeButton"
                ).disabled = false;


            } catch (error) {

                message.textContent =
                    "Upload error: " +
                    error.message;

            }

        }
    );


async function checkOneStudent(student) {

    const username =
        String(
            student.username || ""
        ).trim();


    if (!username) {

        return {
            success: true,
            status: "MISS",
            score: "",
            username: "",
            message:
                "Username is missing"
        };

    }


    try {

        const response =
            await fetch(
                "/check",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            username:
                                username
                        })
                }
            );


        return await response.json();


    } catch (error) {

        return {
            success: false,
            status: "ERROR",
            score: "",
            username: username,
            message:
                "Browser connection error: " +
                error.message
        };

    }

}


document
    .getElementById("testButton")
    .addEventListener(
        "click",
        async function() {

            if (!students.length) {
                return;
            }


            const student =
                students.find(
                    item =>
                        String(
                            item.username || ""
                        ).trim()
                );


            if (!student) {
                return;
            }


            const button =
                document.getElementById(
                    "testButton"
                );


            const message =
                document.getElementById(
                    "actionMessage"
                );


            button.disabled = true;


            message.textContent =
                "Testing one student...";


            const result =
                await checkOneStudent(
                    student
                );


            checkedResults.push(
                result
            );


            renderResults();

            await updateSummary();


            message.textContent =
                "Test completed: " +
                (
                    result.status ||
                    "UNKNOWN"
                );


            button.disabled = false;

        }
    );


async function startChecking(
    resumeMode = false
) {

    if (checking) {
        return;
    }


    checking = true;


    document.getElementById(
        "testButton"
    ).disabled = true;


    document.getElementById(
        "checkAllButton"
    ).disabled = true;


    document.getElementById(
        "resumeButton"
    ).disabled = true;


    const message =
        document.getElementById(
            "actionMessage"
        );


    message.textContent =
        resumeMode
            ? "Resuming checking..."
            : "Checking students...";


    let targetStudents = [];


    if (resumeMode) {

        try {

            const response =
                await fetch("/resume");


            const data =
                await response.json();


            if (data.success) {

                targetStudents =
                    data.students || [];

            }

        } catch (error) {

            message.textContent =
                "Resume error: " +
                error.message;

            checking = false;

            return;

        }

    } else {

        targetStudents =
            students.filter(
                student =>
                    String(
                        student.username || ""
                    ).trim()
            );

    }


    if (!targetStudents.length) {

        message.textContent =
            "No students remaining to check.";

        checking = false;

        document.getElementById(
            "testButton"
        ).disabled = false;

        document.getElementById(
            "checkAllButton"
        ).disabled = false;

        document.getElementById(
            "resumeButton"
        ).disabled = false;

        return;

    }


    const startTime =
        Date.now();


    let processed = 0;


    for (
        let index = 0;
        index < targetStudents.length;
        index++
    ) {

        const student =
            targetStudents[index];


        const result =
            await checkOneStudent(
                student
            );


        checkedResults.push(
            result
        );


        processed++;


        renderResults();

        await updateSummary();


        const elapsed =
            Date.now() - startTime;


        const average =
            elapsed / processed;


        const remaining =
            targetStudents.length -
            processed;


        const remainingMs =
            average * remaining;


        const seconds =
            Math.ceil(
                remainingMs / 1000
            );


        const minutes =
            Math.floor(
                seconds / 60
            );


        const secs =
            seconds % 60;


        document.getElementById(
            "etaText"
        ).textContent =
            "ETA: " +
            minutes +
            "m " +
            secs +
            "s";


        message.textContent =
            "Checking: " +
            processed +
            " / " +
            targetStudents.length;


        if (
            index <
            targetStudents.length - 1
        ) {

            await sleep(500);

        }

    }


    document.getElementById(
        "etaText"
    ).textContent =
        "Checking completed";


    message.textContent =
        resumeMode
            ? "Resume checking completed."
            : "All student checking completed.";


    checking = false;


    document.getElementById(
        "testButton"
    ).disabled = false;


    document.getElementById(
        "checkAllButton"
    ).disabled = false;


    document.getElementById(
        "resumeButton"
    ).disabled = false;


    await updateSummary();

}


document
    .getElementById("checkAllButton")
    .addEventListener(
        "click",
        function() {

            startChecking(false);

        }
    );


document
    .getElementById("resumeButton")
    .addEventListener(
        "click",
        function() {

            startChecking(true);

        }
    );


async function initialize() {

    await loadStudents();

    await loadSavedResults();

    await updateSummary();


    if (students.length) {

        document.getElementById(
            "testButton"
        ).disabled = false;


        document.getElementById(
            "checkAllButton"
        ).disabled = false;


        document.getElementById(
            "resumeButton"
        ).disabled = false;

    }

}


initialize();

</script>

</body>

</html>
