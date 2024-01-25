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
const debug = false;
let header = [
    'Term','Class','Section','Student Name','Student Id',
    'Week Number', 'Assignment Type','Assignment Number', 'Assignment Id', 'Assignment Title',
    'Rubric Id', 'Rubric Name', 'Rubric Line','Line Score','Line Max Score'
].join(',');
header += '\n';

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
    return await getRemainingPagesAsync(url, []);

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
    if(nextLink == null) {
        return listSoFar.concat(responseList);
    } else {
        listSoFar = await getRemainingPagesAsync(nextLink, listSoFar);
        return listSoFar;
    }
}

// escape commas and quotes for CSV formatting
function csvEncode(string) {

    if (typeof(string) === 'undefined') {
        return '';
    }
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

function getItemInModule(contentItem, module) {
    let contentId = contentItem.id;
    let type= 'Assignment';
    if( contentItem.hasOwnProperty('discussion_topic')) {
        type = 'Discussion';
        contentId = contentItem.discussion_topic.id
    }
    if( 'online_quiz' in contentItem.submission_types) {
        type = 'Quiz';
    }
    console.log(type);
    let count = 1;
    for (let item of module.items){
        if (item.type !== type) {
            continue;
        }
        if(item.content_id === contentId) {
            item.numberInModule = count;
            return item;
        }
        count++;
    }
}
function getModuleInfo(contentItem, modules) {
    const regex = /week (\d+)/i;
    let weekCounter = 0;
    for (let module of modules) {
        let match = module.name.match(regex);
        let weekNumber = Number(match? match[1] : null);
        if(!weekNumber) {
            for(let moduleItem in module.items) {
                if(!moduleItem.hasOwnProperty('title')) {
                    continue;
                }
                let match = moduleItem.title.match(regex);
                if (match) {
                    weekNumber = match[1];
                }
            }
        }

        let moduleItem = getItemInModule(contentItem, module);
        if(!moduleItem) {
            continue;
        }
        return {
            weekNumber: weekNumber,
            type: moduleItem.type,
            numberInModule: moduleItem.numberInModule
        }
    }
    return false;
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
    let singleUserSubmissions = userSubmissions.filter(a => a.user_id === user.id);

    //Lets not actually do this if we can't find the user's submissions.
    if (singleUserSubmissions.length === 0) {
        return [];
    }
    let submissions = singleUserSubmissions[0].submissions;
    let skip = false;
    let hide_points, free_form_criterion_comments = false;
    const out = [];


    for (let submission of submissions) {
        let { assignment } = submission;
        const {course_code} = course;
        let section = course_code.match(/-\s*(\d+)$/);
        if (section) {
            section = section[1];
        }
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


        course_code.replace(/^.*_?(\[A-Za-z]{4}\d{3}).*$/, /\1\2/)
        let { weekNumber, numberInModule, type } = getModuleInfo(assignment, modules);
        console.log("number in module", numberInModule);

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
                            if(critKey in critRatingDescs) {
                                crits.rating = critRatingDescs[critKey][critValue.rating_id];
                            } else {
                                console.log('critkey not found ', critKey, critRatingDescs)
                            }
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
                console.log(assignment);
                let rubricName = rubric_settings ? rubric_settings.title : null;
                let name = assignment.hasOwnProperty('name') ? assignment.name : assignment.title;
                let rubricScore = criterion.points;

// let header = [
//     'Term','Class','Section','Student Name','Student Id',
//     'Week Number', 'Assignment Type','Assignment Number', 'Assignment Id',
//     'Rubric Name', 'Rubric Id', 'Rubric Line','Line Score','Line Max Score'
// ].join(',');
                let row = [
                    term.name, course.course_code, section, user.name, user.sis_user_id, weekNumber,
                    type, numberInModule, assignment.id, assignment.name, rubric.description, rubric.id,
                    rubricLine, rubricScore, rubric.points, submission.user.id, submission.user.name,
                ];
                console.log(row);
                row = row.map( item => csvEncode(item));
                let row_string = row.join(',') + '\n';
                out.push(row_string);
            }
        }

    }
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
                let userSubmissions = await getAllPagesAsync(`/api/v1/courses/${courseId}/students/submissions?student_ids=all&per_page=100&include[]=rubric_assessment&include[]=assignment&include[]=user&grouped=true`);
                //let quizzes = await getAllPagesAsync(`/api/v1/courses/${courseId}/quizzes`);

                let accounts = await getAllPagesAsync(`/api/v1/accounts/${course.account_id}`);
                let account = accounts[0];

                let rootAccountId = account.root_account_id;
                let enrollments = await getAllPagesAsync(`/api/v1/courses/${courseId}/enrollments?per_page=100`);
                let modules = await getAllPagesAsync(`/api/v1/courses/${courseId}/modules?include[]=items&include[]=content_details`);

                let response = await fetch(`/api/v1/accounts/${rootAccountId}/terms/${course.enrollment_term_id}`);
                let term = await response.json();


                let rubrics = await getAllPagesAsync(`/api/v1/courses/${courseId}/rubrics?per_page=100`);

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
