// ==UserScript==
// @name         Unity Export Rubric Scores
// @namespace    https://github.com/unity_hallie
// @description  Export all rubric criteria scores for a course to a csv
// @match        https://*/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      1.2
// ==/UserScript==

/* globals $ */
// wait until the window jQuery is loaded
let header = [
    'Term', 'Instructor', 'Class', 'Section', 'Student Name', 'Student Id', 'Enrollment State',
    'Week Number', 'Module', 'Assignment Type', 'Assignment Number', 'Assignment Id', 'Assignment Title',
    'Submission Status', 'Rubric Id', 'Rubric Line', 'Line Name', 'Score', 'Max Score',
].join(',');
header += '\n';


function main() {
    'use strict';
    // utility function for downloading a file
    let saveText = (function () {
        let a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        return function (textArray, fileName, type = 'text') {
            let blob = new Blob(textArray, {type: type}),
                url = window.URL.createObjectURL(blob);
            a.href = url;
            a.download = fileName;
            a.click();
            window.URL.revokeObjectURL(url);
        };
    }());

    $("body").append($('<div id="export_rubric_dialog" title="Export Rubric Scores"></div>'));
    // Only add the export button if a rubric is appearing
    let el = $('#gradebook_header div.statsMetric');
    el.append('<button type="button" class="Button" id="export_one_rubric_btn">Export Assignment</button>');
    el.append('<button type="button" class="Button" id="export_all_rubric_btn">Export All</button>');

    $('#export_one_rubric_btn').click(async function () {
        await exportData(true, saveText);
    });
    $('#export_all_rubric_btn').click(async function () {
        await exportData(false, saveText);
    });

}

async function exportData(singleAssignment = false, saveText) {

    try {
        popUp("Exporting scores, please wait...");
        window.addEventListener("error", showError);


        // Get some initial data from the current URL
        const urlParams = window.location.href.split('?')[1].split('&');
        const courseId = window.location.href.split('/')[4];

        let courseResponse = await fetch(`/api/v1/courses/${courseId}?include=term`)
        let course = await courseResponse.json()

        let accounts = await getAllPagesAsync(`/api/v1/accounts/${course.account_id}`);
        let account = accounts[0];
        let rootAccountId = account.root_account_id;

        const assignId = urlParams.find(i => i.split('=')[0] === "assignment_id").split('=')[1];
        let assignRequest = await fetch(`/api/v1/courses/${courseId}/assignments/${assignId}`);
        let assignment = await assignRequest.json();
        let assignments = await getAllPagesAsync(`/api/v1/courses/${courseId}/assignments`);
        let baseSubmissionsUrl = singleAssignment ? `/api/v1/courses/${courseId}/assignments/${assignId}/submissions` : `/api/v1/courses/${courseId}/students/submissions`;
        let userSubmissions = await getAllPagesAsync(`${baseSubmissionsUrl}?student_ids=all&per_page=100&include[]=rubric_assessment&include[]=assignment&include[]=user&grouped=true`);

        let instructors = await getAllPagesAsync(`/api/v1/courses/${courseId}/users?enrollment_type=teacher`);
        let modules = await getAllPagesAsync(`/api/v1/courses/${courseId}/modules?include[]=items&include[]=content_details`);
        let enrollments = await getAllPagesAsync(`/api/v1/courses/${courseId}/enrollments?per_page=100`);

        let response = await fetch(`/api/v1/accounts/${rootAccountId}/terms/${course.enrollment_term_id}`);
        let term = await response.json();

        let assignmentNames = assignments.map(a => a.name);
        let assignmentsCollection = new AssignmentsCollection(assignments);

        let csvRows = [header];
        for (let enrollment of enrollments) {
            let assignmentId = singleAssignment ? assignId : null;
            let out_rows = await getRows({
                enrollment, modules, userSubmissions, term, course, assignmentId, instructors, assignmentsCollection,
            });
            csvRows = csvRows.concat(out_rows);
        }
        popClose();
        let filename = singleAssignment ? assignment.name : course.course_code;
        saveText(csvRows, `Rubric Scores ${filename.replace(/[^a-zA-Z 0-9]+/g, '')}.csv`);
        saveText([JSON.stringify(userSubmissions, null, 2)], `User Submissions ${filename.replace(/[^a-zA-Z 0-9]+/g, '')}.json`);

        window.removeEventListener("error", showError);

    } catch (e) {
        popClose();
        popUp(`ERROR ${e} while retrieving assignment data from Canvas. Please refresh and try again.`, null);
        window.removeEventListener("error", showError);
        throw (e);
    }
}

/**
 *
 * @param {object} course
 * The course
 * @param {object} enrollment
 * The enrollment of the user to generate rows for
 * @param {array} modules
 * All modules in the course
 * @param {int} assignmentId
 * The ID of the assignment to retrieve data for, if any
 * @param {array} instructors
 * The instructors of the course
 * @param {array} userSubmissions
 * an object containing an array of user submissions { user_id, submissions: []}
 * OR just an array of all users submissions for a single assignment, if assignmentId is specified
 * @param {object} term
 * The term
 * @param {AssignmentsCollection} assignmentsCollection
 * The assignmentsCollection for assignments in this course
 * @returns {Promise<string[]>}
 */
async function getRows(
    {
       course,
       enrollment,
       modules,
       userSubmissions,
       assignmentsCollection,
       instructors,
       term
   }) {
    let {user} = enrollment;
    let singleUserSubmissions = userSubmissions.filter(a => a.user_id === user.id);
    const {course_code} = course;
    let section = course_code.match(/-\s*(\d+)$/);
    let base_code = course_code.match(/([a-zA-Z]{4}\d{3})/);
    if (section) {
        section = section[1];
    }
    if (base_code) {
        base_code = base_code[1]
    }

    let instructorName

    if (instructors.length > 1) {
        instructorName = instructors.map(a => a.name).join(',');
    } else if (instructors.length === 0) {
        instructorName = 'No Instructor Found';
    } else {
        instructorName = instructors[0].name;
    }
    // Let's not actually do this if we can't find the user's submissions.
    if (singleUserSubmissions.length === 0) {
        return [];
    }

    let submissions;
    let entry = singleUserSubmissions[0];
    if (entry.hasOwnProperty('submissions')) {
        submissions = entry.submissions;
    } else {
        submissions = [entry];
    }

    const rows = [];
    let baseRow = [
        term.name,
        instructorName,
        base_code,
        section,

    ]

    for (let submission of submissions) {
        let {assignment} = submission;
        let rubricSettings;

        if (assignment.hasOwnProperty('rubric_settings')) {
            rubricSettings = assignment.rubric_settings;
        }
        let {critOrder, critRatingDescs, critsById} = getCriteriaInfo(assignment);

        course_code.replace(/^.*_?(\[A-Za-z]{4}\d{3}).*$/, /\1\2/)
        let {weekNumber, moduleName, numberInModule, type} = getModuleInfo(assignment, modules, assignmentsCollection);
        let {rubric_assessment: rubricAssessment} = submission;
        let rubricId = typeof (rubricSettings) !== 'undefined' && rubricSettings.hasOwnProperty('id') ?
            rubricSettings.id : 'No Rubric Settings';

        if (user) {
            // Add criteria scores and ratings
            // Need to turn rubric_assessment object into an array
            let critAssessments = []
            let critIds = []

            if (rubricAssessment !== null) {
                for (let critKey in rubricAssessment) {
                    let critValue = rubricAssessment[critKey];
                    let crit = {
                        'id': critKey,
                        'points': critValue.points,
                        'rating': null
                    }
                    if (critValue.rating_id) {
                        if (critKey in critRatingDescs) {
                            crit.rating = critRatingDescs[critKey][critValue.rating_id];
                        } else {
                            console.log('critKey not found ', critKey, critRatingDescs)
                        }
                    }
                    critAssessments.push(crit);
                    critIds.push(critKey);
                }
            }
            let submissionBaseRow = baseRow.concat([
                user.name,
                user.sis_user_id,
                enrollment.enrollment_state,
                weekNumber,
                moduleName,
                type,
                numberInModule,
                assignment.id,
                assignment.name,
                submission.workflow_state,
            ]);

            rows.push(submissionBaseRow.concat([
                rubricId,
                'Total',
                'Total',
                submission.grade,
                assignment.points_possible

            ]))

            // Check for any criteria entries that might be missing; set them to null
            for (let critKey in critOrder) {
                if (!critIds.includes(critKey)) {
                    critAssessments.push({'id': critKey, 'points': null, 'rating': null});
                }
            }
            // Sort into same order as column order
            critAssessments.sort(function (a, b) {
                return critOrder[a.id] - critOrder[b.id];
            });

            for (let critIndex in critAssessments) {
                let critAssessment = critAssessments[critIndex];
                let criterion = critsById[critAssessment.id];

                rows.push(submissionBaseRow.concat([
                    criterion.id,
                    Number(critIndex) + 1,
                    criterion.description,
                    critAssessment.points,
                    criterion.points
                ]));
            }
        }
    }

    let out = [];
    for (let row of rows) {
        let row_string = row.map(item => csvEncode(item)).join(',') + '\n';
        out.push(row_string);
    }
    return out;
}

// escape commas and quotes for CSV formatting
function csvEncode(string) {

    if (typeof (string) === 'undefined' || string === null || string === 'null') {
        return '';
    }
    string = String(string);

    if (string) {
        string = string.replace(/(")/g, '"$1');
        string = string.replace(/\s*\n\s*/g, ' ');
    }
    return `"${string}"`;
}

function showError(event) {
    popUp(event.message);
    window.removeEventListener("error", showError);
}

function getModuleInfo(contentItem, modules, assignmentsById) {
    const regex = /(week|module) (\d+)/i;

    for (let module of modules) {
        let match = module.name.match(regex);
        let weekNumber = !match? null : Number(match[1]);
        if (!weekNumber) {
            for (let moduleItem of module.items) {
                if (!moduleItem.hasOwnProperty('title')) {
                    continue;
                }
                let match = moduleItem.title.match(regex);
                if (match) {
                    weekNumber = match[2];
                }
            }
        }

        let moduleItem = getItemInModule(contentItem, module, assignmentsById);
        if (!moduleItem) {
            continue;
        }
        return {
            weekNumber: weekNumber == null? '-' : weekNumber,
            moduleName: module.name,
            type: moduleItem.type,
            numberInModule: moduleItem.numberInModule
        }
    }
    return false;
}


function getItemInModule(contentItem, module, assignmentsCollection) {

    let contentId;
    let type = assignmentsCollection.getAssignmentContentType(contentItem);
    if (type === 'Discussion') {
        contentId = contentItem.discussion_topic.id;
    } else if (type === 'Quiz') {
        contentId = contentItem.quiz_id;
    } else {
        contentId = contentItem.id;
    }

    let count = 1;
    for (let moduleItem of module.items) {

        let moduleItemAssignment = assignmentsCollection.getContentById(moduleItem.content_id);
        if (assignmentsCollection.getModuleItemType(moduleItem) !== type){
          continue;
        }

        if (moduleItem.content_id === contentId) {
            if (type === 'Discussion' && !contentItem.hasOwnProperty('rubric')) {
                moduleItem.numberInModule = '-';
            } else {
                moduleItem.numberInModule = count;
            }
            moduleItem.type = type;
            return moduleItem;
        }

        if (type === 'Discussion' && !moduleItemAssignment.hasOwnProperty('rubric')){
            continue;
        }

        count++;
    }
}

/**
 * Fill out the csv header and map criterion ids to sort index
 * Also create an object that maps criterion ids to an object mapping rating ids to descriptions
 * @param assignment
 * The assignment from canvas api
 * @returns {{critRatingDescs: *[], critsById: *[], critOrder: *[]}}
 */
function getCriteriaInfo(assignment) {
    if (!assignment || !assignment.hasOwnProperty('rubric')) {
        return {critsById: [], critRatingDescs: [], critOrder: []}
    }
    let rubricCriteria = assignment.rubric;

    let critOrder = [];
    let critRatingDescs = [];
    let critsById = [];
    for (let critIndex in rubricCriteria) {
        let rubricCriterion = rubricCriteria[critIndex];
        critOrder[rubricCriterion.id] = critIndex;
        critRatingDescs[rubricCriterion.id] = {};
        critsById[rubricCriterion.id] = rubricCriterion;

        for (let rating of rubricCriterion.ratings) {
            critRatingDescs[rubricCriterion.id][rating.id] = rating.description;
        }
    }
    return {critOrder, critRatingDescs, critsById}
}

function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    } else {
        setTimeout(async function () {
            defer(method);
        }, 100);
    }
}

function popUp(text) {
    let el = $("#export_rubric_dialog");
    el.html(`<p>${text}</p>`);
    el.dialog({buttons: {}});
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
    if (headers.has('link')) {
        let linkStr = headers.get('link');
        let links = linkStr.split(',');
        nextLink = null;
        for (let link of links) {
            if (link.split(';')[1].includes('rel="next"')) {
                nextLink = link.split(';')[0].slice(1, -1);
            }
        }
    }
    if (nextLink == null) {
        return listSoFar.concat(responseList);
    } else {
        return await getRemainingPagesAsync(nextLink, listSoFar.concat(responseList));
    }
}

/**
 * A collection of assignments grabbed from the submissions that returns and finds them in various ways
 */
class AssignmentsCollection {
    constructor(assignments) {
        this.assignments = assignments;

        this.assignmentsById = {}
        for (let assignment of assignments) {
            this.assignmentsById[assignment.id] = assignment;
        }

        this.discussions = assignments.filter(assignment => assignment.hasOwnProperty('discussion_topic'))
            .map(function (assignment) {
                let discussion = assignment.discussion_topic;
                discussion.assignment = assignment;
                return discussion;
            });

        this.discussionsById = {};
        this.assignmentsByDiscussionId = {};
        for (let discussion of this.discussions) {
            this.discussionsById[discussion.id] = discussion;
            this.assignmentsByDiscussionId[discussion.id] = discussion.assignment;

        }

        this.assignmentsByQuizId = {};
        for (let assignment of assignments.filter(a => a.hasOwnProperty('quiz_id'))) {
            this.assignmentsByQuizId[assignment.quiz_id] = assignment;
        }
    }

    /**
     * Gets content by id
     * @param id the primary id of that content item (not necessarily the assignment Id)
     * The content_id property that it would have were it in a module
     * @returns {*}
     */
    getContentById(id) {
        for (let collection of [
            this.assignmentsByQuizId,
            this.assignmentsByDiscussionId,
            this.assignmentsById
        ]) {
            if (collection.hasOwnProperty(id)) {
                return collection[id];
            }
        }
    }



    /**
     * Returns content type as a string if it is an Assignment, Quiz, or Discussion
     * @param contentItem
     * the content item
     * @returns {string}
     */
    getAssignmentContentType(contentItem) {
        if(contentItem.hasOwnProperty('submission_types')) {
            if (contentItem.submission_types.includes('external_tool')) { return 'External Tool'}
        }
        if(contentItem.hasOwnProperty('discussion_topic')) { return 'Discussion'}
        if(contentItem.hasOwnProperty('quiz_id')) { return 'Quiz'}
        let id = contentItem.hasOwnProperty('id') ? contentItem.id : contentItem;

        if (this.assignmentsByQuizId.hasOwnProperty(id)) {
            return "Quiz";
        } else if (this.assignmentsByDiscussionId.hasOwnProperty(id)) {
            return 'Discussion';
        } else {
            return 'Assignment';
        }
    }

    getModuleItemType(moduleItem) {
        console.log(moduleItem);
        if (moduleItem.type !== 'Assignment') return moduleItem.type;
        const assignment = this.assignmentsById[moduleItem.content_id];
        console.log(moduleItem, this.getAssignmentContentType(assignment))
        return this.getAssignmentContentType(assignment);
    }
}


defer(main);
