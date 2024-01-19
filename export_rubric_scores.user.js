// ==UserScript==
// @name         Export Rubric Scores
// @namespace    https://github.com/unity_hallie
// @description  Export all rubric criteria scores for an assignment to a CSV
// @match        https://*/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      0.1
// ==/UserScript==

/* globals $ */

// wait until the window jQuery is loaded
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
    $("#export_rubric_dialog").html(`<p>${text}</p>`);
    $("#export_rubric_dialog").dialog({ buttons: {} });
}

function popClose() {
    $("#export_rubric_dialog").dialog("close");
}

function getAllPages(url, callback) {
   return getRemainingPages(url, [], callback);
}

// Recursively work through paginated JSON list
function getRemainingPages(nextUrl, listSoFar, callback) {
    return $.getJSON(nextUrl, function(responseList, textStatus, jqXHR) {
        var nextLink = null;
        $.each(jqXHR.getResponseHeader("link").split(','), function (linkIndex, linkEntry) {
            if (linkEntry.split(';')[1].includes('rel="next"')) {
                nextLink = linkEntry.split(';')[0].slice(1, -1);
            }
        });
        if (nextLink == null) {
            // all pages have been retrieved
            callback(listSoFar.concat(responseList));
        } else {
            getRemainingPages(nextLink, listSoFar.concat(responseList), callback);
        }
    }).fail(function (jqXHR, textStatus, errorThrown) {
        popUp(`ERROR ${jqXHR.status} while retrieving data from Canvas. Url: ${nextUrl}<br/><br/>Please refresh and try again.`, null);
        window.removeEventListener("error", showError);
    });
}


async function getAllPagesAsync(url) {
    let out = await getRemainingPagesAsync(url, []);
    console.log("Get All Pages", out);
    return out;
}

async function getRemainingPagesAsync(url, listSoFar) {
    let response = await fetch(url);
    let responseList = await response.json();
console.log(response);
    let headers = response.headers;

    let nextLink;
    if(headers.has('link')){
        let linkStr = headers.get('link');
        console.log("LinkStr", linkStr);
        let links = linkStr.split(',');
        nextLink = null;
        for(let link of links){
            if (link.split(';')[1].includes('rel="next"')) {
                nextLink = link.split(';')[0].slice(1, -1);
            }
        }
    }
    if(nextLink == null){
        console.log('listSoFar1', listSoFar)
        return listSoFar.concat(responseList);
    } else {
        listSoFar = await getRemainingPagesAsync(nextLink, listSoFar);
        console.log("listSoFar2", listSoFar);
        return listSoFar;
    }
}

// escape commas and quotes for CSV formatting
function csvEncode(string) {
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

defer(function() {
    'use strict';
console.log("RUn2");
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
        $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="export_rubric_btn">Export Rubric Scores</button>');
        $('#export_rubric_btn').click(async function() {
            try{
                popUp("Exporting scores, please wait...");
                window.addEventListener("error", showError);

                // Get some initial data from the current URL
                const courseId = window.location.href.split('/')[4];
                const urlParams = window.location.href.split('?')[1].split('&');
                const assignId = urlParams.find(i => i.split('=')[0] === "assignment_id").split('=')[1];

                let assignments = await getAllPagesAsync(`/api/v1/courses/${courseId}/assignments?per_page=100`);
                let enrollments = await getAllPagesAsync(`/api/v1/courses/${courseId}/enrollments?per_page=100`);
                //let user_submissions = await getAllPagesAsync(`/api/v1/courses/${courseId}/students/submissions?per_page=100&include[]=user&include=rubric_assessment&group=True`);
                let submissions = await getAllPagesAsync(`/api/v1/courses/${courseId}/assignments/${assignId}/submissions?include[]=rubric_assessment&per_page=100`);

                console.log(submissions);
                let assignment = await $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`);
                // Get the rubric data

                if (!('rubric_settings' in assignment)) {
                    popUp(`ERROR: No rubric settings found at /api/v1/courses/${courseId}/assignments/${assignId}.<br/><br/> `
                        + 'This is likely due to a Canvas bug where a rubric has entered a "soft-deleted" state. '
                        + 'Please use the <a href="https://community.canvaslms.com/t5/Canvas-Admin-Blog/Undeleting-things-in-Canvas/ba-p/267116">Undelete feature</a> '
                        + 'to restore the rubric associated with this assignment or contact Canvas Support.');
                    return;
                }
                // If rubric is set to hide points, then also hide points in export
                // If rubric is set to use free form comments, then also hide ratings in export
                const hidePoints = assignment.rubric_settings.hide_points;
                const hideRatings = assignment.rubric_settings.free_form_criterion_comments;
                if (hidePoints && hideRatings) {
                    popUp("ERROR: This rubric is configured to use free-form comments instead of ratings AND to hide points, so there is nothing to export!");
                    return;
                }

                // Fill out the csv header and map criterion ids to sort index
                // Also create an object that maps criterion ids to an object mapping rating ids to descriptions
                var critOrder = {};
                var critRatingDescs = {};
                var header = "Student Name,Student ID,Posted Score,Attempt Number";
                $.each(assignment.rubric, function (critIndex, criterion) {
                    critOrder[criterion.id] = critIndex;
                    if (!hideRatings) {
                        critRatingDescs[criterion.id] = {};
                        $.each(criterion.ratings, function (i, rating) {
                            critRatingDescs[criterion.id][rating.id] = rating.description;
                        });
                        header += ',' + csvEncode('Rating: ' + criterion.description);
                    }
                    if (!hidePoints) {
                        header += ',' + csvEncode('Points: ' + criterion.description);
                    }
                });
                header += '\n';

                // Iterate through submissions
                var csvRows = [header];
                $.each(submissions, function (subIndex, submission) {
                    const user = enrollments.find(i => i.user_id === submission.user_id).user;
                    if (user) {
                        var row = `${user.name},${user.sis_user_id},${submission.score},${submission.attempt}`;
                        // Add criteria scores and ratings
                        // Need to turn rubric_assessment object into an array
                        var crits = []
                        var critIds = []
                        if (submission.rubric_assessment != null) {
                            $.each(submission.rubric_assessment, function (critKey, critValue) {
                                if (hideRatings) {
                                    crits.push({'id': critKey, 'points': critValue.points, 'rating': null});
                                } else {
                                    crits.push({
                                        'id': critKey,
                                        'points': critValue.points,
                                        'rating': critRatingDescs[critKey][critValue.rating_id]
                                    });
                                }
                                critIds.push(critKey);
                            });
                        }
                        // Check for any criteria entries that might be missing; set them to null
                        $.each(critOrder, function (critKey, critValue) {
                            if (!critIds.includes(critKey)) {
                                crits.push({'id': critKey, 'points': null, 'rating': null});
                            }
                        });
                        // Sort into same order as column order
                        crits.sort(function (a, b) {
                            return critOrder[a.id] - critOrder[b.id];
                        });
                        $.each(crits, function (critIndex, criterion) {
                            if (!hideRatings) {
                                row += `,${csvEncode(criterion.rating)}`;
                            }
                            if (!hidePoints) {
                                row += `,${criterion.points}`;
                            }
                        });
                        row += '\n';
                        csvRows.push(row);
                    }
                });
                popClose();
                saveText(csvRows, `Rubric Scores ${assignment.name.replace(/[^a-zA-Z 0-9]+/g, '')}.csv`);
                window.removeEventListener("error", showError);
            } catch(e) {
                popClose();
                popUp(`ERROR ${e} while retrieving assignment data from Canvas. Please refresh and try again.`, null);
                throw(e);
                window.removeEventListener("error", showError);
            }
        });
    }
});
