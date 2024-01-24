// ==UserScript==
// @name         Unity Export Rubric Scores
// @namespace    https://github.com/unity_hallie
// @description  Export all rubric criteria scores for a course to a csv
// @match        https://*/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      0.3
// ==/UserScript==

/* globals $ */
// wait until the window jQuery is loaded
const debug = false
function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    }
    else {
        setTimeout(async function() { defer(method); }, 100);
    }
}

function waitForElement(selector, callback) {
    if ($(selector).length) {
        callback();
    } else {
        setTimeout(function() {
            waitForElement(selector, callback);
        }, 100);
    }
}

function popUp(text) {
    let el = $("#export_rubric_dialog");
    el.html(`<p>${text}</p>`);
    el.dialog({ buttons: {} });
}

function popClose() {
    $("#export_rubric_dialog").dialog("close");
}

async function getAllPagesAsync(url) {
    let out = await getRemainingPagesAsync(url, []);
    return out;
}

async function getRemainingPagesAsync(url, listSoFar) {
    let response = await fetch(url);
    let responseList = await response.json();
    let headers = response.headers;

    let nextLink;
    if(headers.has('link')){
        let linkStr = headers.get('link');
        let links = linkStr.split(',');
        nextLink = null;
        for(let link of links){
            if (link.split(';')[1].includes('rel="next"')) {
                nextLink = link.split(';')[0].slice(1, -1);
            }
        }
    }
    if(nextLink == null || debug){
        return listSoFar.concat(responseList);
    } else {
        listSoFar = await getRemainingPagesAsync(nextLink, listSoFar);
        return listSoFar;
    }
}

// escape commas and quotes for CSV formatting
function csvEncode(string) {
    string = String(string);
    if(string) {
        string = string.replace(/(")/g,'"$1');
        string = string.replace(/\s*\n\s*/g,' ');
    }
    return string;
}

function showError(event) {
    popUp(event.message);
    window.removeEventListener("error", showError);
}

/**
 *
 * @param {object} course
 * @param {object} enrollment
 * @param {array} modules
 * @param {array} quizzes
 * @param {array} userSubmissions
 * @param {object} term
 * @returns {Promise<string[]>}
 */
async function getEnrollmentRows({ course, enrollment, modules, quizzes, userSubmissions,
                                     rubrics, term }){
    let { user } = enrollment;
    let submissions = userSubmissions.filter(a => a.user_id === user.id);
    let skip = false;
    let hide_points, free_form_criterion_comments = false
    const out = [];



    for (let submission of submissions) {
        let { assignment } = submission;
        let { rubric } = assignment;
        if (!('rubric_settings' in assignment)) {
            skip = true;

        }

        if(!assignment.rubric_settings) {
            skip = true;
        } else {
            let rs = assignment.rubric_settings;
            hide_points = rs.hide_points;
            free_form_criterion_comments = rs.free_form_criterion_comments;
            if (hide_points && free_form_criterion_comments) {
                popUp(`ERROR: ${assignment.name} is configured to use free-form comments instead of ratings AND to hide points`
                    +`, so there is nothing to export!`);
                skip = true;
            }
        }

        // Fill out the csv header and map criterion ids to sort index
        // Also create an object that maps criterion ids to an object mapping rating ids to descriptions
        let critOrder = {};
        let critRatingDescs = {};
        for (let critIndex in rubric){
            let criterion = rubric[critIndex];
            critOrder[criterion.id] = critIndex;
            critRatingDescs[criterion.id] = {};
            for(let rating of criterion.ratings) {
                critRatingDescs[criterion.id][rating.id] = rating.description;
            }
        }

        // Iterate through submissions
        for (let submission of submissions) {
            const {assignment} = submission;
            const {user} = enrollment;
            const {course_code} = course;
            let section = course_code.match(/-\s*(\d+)$/);
            if (section) {
                section = section[1];
            }
            console.log(JSON.stringify(term));
            course_code.replace(/^.*_?(\[A-Za-z]{4}\d{3}).*$/, /\1\2/)
            let baseEntry = {
                assignmentTotalScore: submission.score,
                assignmentType: assignment.submission_types,
                assignmentNumber: assignment.title,
                assignmentId: assignment.id,
                attemptNumber: submission.attempt,
                courseCode: course["course_code"],
                rubricLineMaxScore: null,
                rubricLineNumber: null,
                rubricLineScore: null,
                section: section,
                studentId: user.sis_user_id,
                studentName: user.name,
                term: term,
                weekNumber: null,
            }

            if (user) {
                // Add criteria scores and ratings
                // Need to turn rubric_assessment object into an array
                let crits = []
                let critIds = []
                let { rubric_assessment } = submission;
                if (rubric_assessment == null) {

                } else {
                    for (let critKey in rubric_assessment) {
                        let critValue = rubric_assessment[critKey];
                        if (free_form_criterion_comments) {
                            crits.push({'id': critKey, 'points': critValue.points, 'rating': null});
                        } else {
                            let crit = {
                                'id': critKey,
                                'points': critValue.points,
                            }
                            if(critValue.rating_id) {
                                crits.rating = critRatingDescs[critKey][critValue.rating_id]
                            }
                        }
                        critIds.push(critKey);
                    }
                }
                // Check for any criteria entries that might be missing; set them to null
                for (let critKey in critOrder) {
                    if (!critIds.includes(critKey)) {
                        crits.push({'id': critKey, 'points': null, 'rating': null});
                    }
                }
                // Sort into same order as column order
                crits.sort(function (a, b) {
                    return critOrder[a.id] - critOrder[b.id];
                });

                for(let critIndex in crits) {
                    let criterion = crits[critIndex];
                    let { rubric_settings } = assignment;
                    let rubricLine = critIndex;
                    let rubricName = rubric_settings ? rubric_settings.title : null;
                    let assignmentNumber = 0;
                    let weekNumber = 0;
                    let row = [
                        baseEntry.term.name, baseEntry.courseCode, baseEntry.section, baseEntry.studentName,
                        baseEntry.studentId, baseEntry.weekNumber, baseEntry.assignmentType,
                        assignmentNumber, baseEntry.assignmentId, rubricName,
                        rubricLine, criterion.points, rubric.id,
                        criterion.rating, JSON.stringify(criterion)
                    ];

                    row = row.map( item => csvEncode(item));
                    let row_string = row.join(',') + '\n';
                    out.push(row_string);
                }
            }
        }

    }
    console.log(out);
    return out;

}


defer(function() {
    'use strict';
    // utility function for downloading a file
    let saveText = (function () {
        var a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        return function (textArray, fileName) {
            var blob = new Blob(textArray, {type: "text"}),
                url = window.URL.createObjectURL(blob);
            a.href = url;
            a.download = fileName;
            a.click();
            window.URL.revokeObjectURL(url);
        };
    }());

    $("body").append($('<div id="export_rubric_dialog" title="Export Rubric Scores"></div>'));
    // Only add the export button if a rubric is appearing
    if ($('#rubric_summary_holder').length > 0) {
        $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="export_all_rubric_btn">Export All</button>');
        $('#export_all_rubric_btn').click(async function() {
            try{
                popUp("Exporting scores, please wait...");
                window.addEventListener("error", showError);

               // Get some initial data from the current URL
                const urlParams = window.location.href.split('?')[1].split('&');
                const courseId = window.location.href.split('/')[4];

                let courseResponse = await fetch(`/api/v1/courses/${courseId}?include=term`)
                let course = await courseResponse.json()

                let assignments = await getAllPagesAsync(`/api/v1/courses/${courseId}/assignments?per_page=100`);
                let userSubmissions = await getAllPagesAsync(`/api/v1/courses/${courseId}/students/submissions?student_ids=all&per_page=100&include[]=rubric_assessment&include[]=assignment&group=True`);
                //let quizzes = await getAllPagesAsync(`/api/v1/courses/${courseId}/quizzes`);

                let account = await getAllPagesAsync(`/api/v1/accounts/${course.account_id}`);
                let rootAccountId = account.root_account_id;
                let enrollments = await getAllPagesAsync(`/api/v1/courses/${courseId}/enrollments?per_page=100`);
                let modules = await getAllPagesAsync(`/api/v1/courses/${courseId}/modules`);

                let termResponse = await fetch(`/api/v1/accounts/${rootAccountId}/terms/${course.enrollment_term_id}`)
                let term = await termResponse.json()

                let rubrics = await getAllPagesAsync(`/api/v1/courses/${courseId}/rubrics?per_page=100`);

                let header = [
                    'Term','Class','Section','Student Name','Student Id',
                    'Week Number', 'Assignment Type','Assignment Number',
                    'Rubric Line','Line Score','Line Max Score', 'X', 'Y', 'Z'
                ].join(',');

                header += '\n';
                let csvRows = [header];
                for(let enrollment of enrollments) {
                    let out_rows = await getEnrollmentRows({
                        enrollment, modules, assignments, userSubmissions, term, rubrics, course,
                    });
                    csvRows = csvRows.concat(out_rows);
                }
                popClose();
                console.log(csvRows);
                saveText(csvRows, `Rubric Scores ${course.course_code.replace(/[^a-zA-Z 0-9]+/g, '')}.csv`);
                window.removeEventListener("error", showError);
            } catch(e) {
                popClose();
                popUp(`ERROR ${e} while retrieving assignment data from Canvas. Please refresh and try again.`, null);
                window.removeEventListener("error", showError);
                throw(e);
            }
        });
    }
});
