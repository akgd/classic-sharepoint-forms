// closest polyfill
if (!Element.prototype.matches)
    Element.prototype.matches = Element.prototype.msMatchesSelector ||
        Element.prototype.webkitMatchesSelector;

if (!Element.prototype.closest)
    Element.prototype.closest = function (s) {
        var el = this;
        if (!document.documentElement.contains(el)) return null;
        do {
            if (el.matches(s)) return el;
            el = el.parentElement || el.parentNode;
        } while (el !== null && el.nodeType === 1);
        return null;
    };

// Check item form view
function checkFormView(currentUrl) {
    var classicWorkspace = document.querySelector('body');

    var newForm = checkForSubstring(currentUrl, 'newform');
    var dispForm = checkForSubstring(currentUrl, 'dispform');
    var editForm = checkForSubstring(currentUrl, 'editform');

    if (newForm) {
        classicWorkspace.classList.add('cspf-new-form');
        return 'cspf-new-form';
    } else if (editForm) {
        classicWorkspace.classList.add('cspf-edit-form');
        return 'cspf-edit-form';
    } else if (dispForm) {
        classicWorkspace.classList.add('cspf-disp-form');
        return 'cspf-disp-form';
    } else {
        console.log('Not a valid form URL.')
    }
}

// Add IDs to table based on the view type
function addRowIDs(view) {
    var workspace = document.getElementById('s4-workspace');
    workspace.classList.add('classic-cspf-form');

    hideAllEls(document.querySelectorAll('.ms-formtoolbar')); // hide sp form table
    hideAllEls(document.querySelectorAll('.ms-formtable')); // hide sp buttons and meta

    if ((view === 'cspf-new-form') || (view === 'cspf-edit-form')) {
        // Editable view, target header elements to find internal field name
        var findFields = document.querySelectorAll('.ms-standardheader');
        for (var i = 0; i < findFields.length; i++) {
            var getId = findFields[i].id;
            var fieldEl = document.getElementById(getId);
            var row = fieldEl.closest('tr');
            row.id = 'cspf-' + getId;
        }
    } else if (view === 'cspf-disp-form') {
        // Display view, tarket a elements to find internal field name
        var findFields = document.querySelectorAll('a[name^="SPBookmark_"]');
        for (var i = 0; i < findFields.length; i++) {
            var getName = findFields[i].getAttribute('name');
            var getId = getName.replace('SPBookmark_', '');
            var row = findFields[i].closest('tr');
            row.id = 'cspf-' + getId;
        }
    }
}

// To generate a form section from an array
// Handles fields in classic edit and display forms
/* Example of expected arr:
var yourFields = [
    {
        internalName: 'Title',
        displayName: 'Project Title',
        helpText: '',
        helpAfterField: true,
        specialClasses: ''
    },
    {
        internalName: 'ProjLead',
        displayName: 'Project Lead',
        helpText: '',
        helpAfterField: true,
        specialClasses: ''
    }
]*/
function generateFormSection(arr, targetElId) {
    var currentView = checkFormView(window.location.href);
    var workspace = document.getElementById('s4-workspace');
    var checkIfAddedRowIDs = elHasClass(workspace, 'classic-cspf-form');
    if (!checkIfAddedRowIDs) {
        addRowIDs(currentView);
    }

    for (i = 0; i < arr.length; i++) {
        var intName = arr[i].internalName;
        var specialClasses = arr[i].specialClasses;
        var newRow = '#cspf-field-' + intName;

        var newEl = '<div id="cspf-field-' + intName + '" class="cspf-field-row ' + specialClasses + '"><div class="cspf-field-title">' + arr[i].displayName + '</div><div class="cspf-field-body"><div id="cspf-field-data-' + intName + '" class="cspf-field-data"></div><div class="cspf-field-error help is-danger" style="display:none">Please complete this field.</div></div></div>';

        // handle help text
        if (arr[i].help !== '') {
            if (arr[i].helpAfterField === true) {
                newEl = '<div id="cspf-field-' + intName + '" class="cspf-field-row ' + specialClasses + '"><div class="cspf-field-title">' + arr[i].displayName + '</div><div class="cspf-field-body"><div id="cspf-field-data-' + intName + '" class="cspf-field-data"></div><div class="cspf-field-help">' + arr[i].helpText + '</div><div class="cspf-field-error help is-danger" style="display:none">Please complete this field.</div></div></div>';
            } else {
                newEl = '<div id="cspf-field-' + intName + '" class="cspf-field-row ' + specialClasses + '"><div class="cspf-field-title">' + arr[i].displayName + '</div><div class="cspf-field-body"><div class="cspf-field-help">' + arr[i].helpText + '</div><div id="cspf-field-data-' + intName + '" class="cspf-field-data"></div><div class="cspf-field-error help is-danger" style="display:none">Please complete this field.</div></div></div>';
            }
        }

        //Insert new element into target div
        document.querySelector('#' + targetElId).insertAdjacentHTML('beforeend', newEl);

        if ((currentView === 'cspf-new-form') || (currentView === 'cspf-edit-form')) {
            //Find the actual SP form data field
            var spFieldRowId = document.getElementById('cspf-' + arr[i].internalName);
            var spFieldBody = spFieldRowId.getElementsByClassName('ms-formbody');
            spFieldBody[0].removeAttribute('width');
            var spFieldEl = spFieldBody[0];
            document.querySelector(newRow + ' .cspf-field-data').appendChild(spFieldEl);
        } else if (currentView === 'cspf-disp-form') {
            var targetFieldSel = document.getElementById('cspf-' + arr[i].internalName);
            var targetFieldBody = targetFieldSel.getElementsByClassName('ms-formbody');
            var targetField = targetFieldBody[0];
            if (targetField) {
                var targetFieldContents = '<div class="cspf-wrap">' + targetField.innerHTML + '</div>';
                document.querySelector(newRow + ' .cspf-field-data').innerHTML = targetFieldContents;
            } else {
                console.log(targetFieldSel)
            }
        }
    }

    if (currentView === 'cspf-new-form') {
        hideByClass('cspf-hide-on-new');
    } else if (currentView === 'cspf-edit-form') {
        hideByClass('cspf-hide-on-edit');
    } else if (currentView === 'cspf-disp-form') {
        hideByClass('cspf-hide-on-disp');
    }
}
// help
function hideByClass(name) {
    var toHide = document.getElementsByClassName(name); //toHide is an array
    for (var i = 0; i < toHide.length; i++) {
        toHide[i].style.display = 'none';
    }
}

// Get URL parameter value
function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

// Change URL parameter
function replaceUrlParam(url, paramName, paramValue) {
    if (paramValue == null) {
        paramValue = '';
    }
    var pattern = new RegExp('\\b(' + paramName + '=).*?(&|#|$)');
    if (url.search(pattern) >= 0) {
        return url.replace(pattern, '$1' + paramValue + '$2');
    }
    url = url.replace(/[?#]$/, '');
    return url + (url.indexOf('?') > 0 ? '&' : '?') + paramName + '=' + paramValue;
}


// Get people picker entries
function getPeoplePicker(internalName) {
    var picker = document.querySelector('#cspf-field-data-' + internalName + ' input[id$="_HiddenInput"]');
    if (picker.value) {
        var data = JSON.parse(picker.value);
        return data;
    } else {
        // picker is empty
    }
}

// Handle conditional fields (selectors as arrays)
function handleConditionals(internalName, inputValue, showEls, hideEls) {
    var hideEls = hideEls || [];
    var el;

    var checkRadio = document.querySelector('#cspf-field-' + internalName + ' input:checked');
    var multiChoiceEl = document.querySelector('#cspf-field-' + internalName + ' .cspf-field-data table[id$="MultiChoiceTable"]');
    var checkInput = document.querySelector('#cspf-field-' + internalName + ' input');
    var checkSelect = document.querySelector('#cspf-field-' + internalName + ' select');
    var checkDispEl = document.querySelector('#cspf-field-' + internalName + ' .cspf-field-data');

    if (multiChoiceEl) {
        var checkboxes = document.querySelectorAll('#cspf-field-' + internalName + ' .cspf-field-data table[id$="MultiChoiceTable"] input[type=checkbox]');

        for (var i = 0; i < checkboxes.length; i++) {
            checkMultiChoiceVal(checkboxes[i], inputValue, showEls, hideEls);
            checkboxes[i].addEventListener('change', function (event) {
                checkMultiChoiceVal(event.target, inputValue, showEls, hideEls);
            });
        }
    } else if (checkRadio) {
        var checkedRadio = document.querySelector('#cspf-field-' + internalName + ' input:checked');

        if (checkedRadio) {
            showHideIfValue(checkedRadio, inputValue, showEls, hideEls);
        } else {
            hideAllEls(showEls)
        }

        var fieldRadios = document.querySelectorAll('#cspf-field-' + internalName + ' input');
        for (i=0; i < fieldRadios.length; i++) {
            fieldRadios[i].addEventListener('change', function (event) {
                var radioEL = event.target;
                var radioChecked = this.checked
                if (radioChecked) {
                    showHideIfValue(this, inputValue, showEls, hideEls);
                }
            });
        }
        
    } else {
        if (checkSelect) {
            var el = checkSelect;
        } else if (checkInput) {
            var el = checkInput;
        } else if (checkDispEl) {
            var el = checkDispEl;
        } else {
            console.log('Not a valid field type');
            // invalid
        }
        if (el) {
            showHideIfValue(el, inputValue, showEls, hideEls);

            // Specify events we want to listen for
            var eventTypes = ['input', 'change'];

            // Add listeners per event type
            for (var e = 0; e < eventTypes.length; e++) {
                el.addEventListener(eventTypes[e], function (event) {
                    if (el) {
                        showHideIfValue(el, inputValue, showEls, hideEls)
                    }

                });
            }
        }
    }
}

function checkMultiChoiceVal(target, val, showEls, hideEls) {
    var isChecked = target.checked;

    var labels = document.getElementsByTagName('label');

    for (var i = 0; i < labels.length; i++) {
        if (labels[i].htmlFor != '') {
            var elem = document.getElementById(labels[i].htmlFor);
            if (elem)
                elem.label = labels[i];
        }
    }

    var labelVal = document.getElementById(target.id).label.innerText;

    if (val === labelVal) {
        if (isChecked) {
            showHideEls(showEls, hideEls);
        } else {
            showHideEls(hideEls, showEls);
        }
    }
}

function getLabelsForInputElement(element) {
    var labels = [];
    var id = element.id;

    if (element.labels) {
        return element.labels;
    }

    id && Array.prototype.push
        .apply(labels, document.querySelector("label[for='" + id + "']"));

    while (element = element.parentNode) {
        if (element.tagName.toLowerCase() == "label") {
            labels.push(element);
        }
    }
    return labels;
};

function showHideEls(showEls, hideEls) {
    // Show elements based on selector
    for (var i = 0; i < showEls.length; i++) {
        var getEls = document.querySelectorAll(showEls[i]);
        for (el=0; el < getEls.length; el++) {
            getEls[el].style.display = '';
        }
    }

    // Hide elements based on selector
    for (var i = 0; i < hideEls.length; i++) {
        var getEls = document.querySelectorAll(hideEls[i]);
        for (el=0; el < getEls.length; el++) {
            getEls[el].style.display = 'none';
        }
    }
}

// Show and hide elements based on a specified value
function showHideIfValue(el, inputValue, showEls, hideEls) {
    
    if (el.value) {
        var getValue = el.value;
    } else {
        var getValue = el.textContent;
        getValue = getValue.trim();
    }

    var inputToString = inputValue.toString();
    var getValToString = getValue.toString();

    if (inputToString === getValToString) {
        // Show elements based on selector
        for (var i = 0; i < showEls.length; i++) {
            var getEls = document.querySelectorAll(showEls[i]);
            for (el=0; el < getEls.length; el++) {
                getEls[el].style.display = '';
            }
        }
        // Hide elements based on selector
        for (var i = 0; i < hideEls.length; i++) {
            var getEls = document.querySelectorAll(hideEls[i]);
            for (el=0; el < getEls.length; el++) {
                getEls[el].style.display = 'none';
            }
        }
    } else {
        // The selected value isn't what we are looking for
        // Hide elements based on selector
        for (var i = 0; i < showEls.length; i++) {
            var getEls = document.querySelectorAll(showEls[i]);
            for (el=0; el < getEls.length; el++) {
                getEls[el].style.display = 'none';
            }
        }
    }

}

// hiding els
function hideElement(el) {
    el.style.display = 'none';
}

function hideAllEls(elsToHide) {
    for (var i = 0; i < elsToHide.length; i++) {
        elsToHide[i].style.visibility = "hidden"; // or
        elsToHide[i].style.display = "none"; // depending on what you're doing
    }
}
// to get choice field value
function getFieldValue(internalName) {
    var el = '#cspf-field-' + internalName + ' [id^="' + internalName + '_"]';
    return document.querySelector(el).value;
}

// Set choice
function setFieldValue(internalName, val) {
    var el = '#cspf-field-' + internalName + ' [id^="' + internalName + '_"]';
    document.querySelector(el).value = val;
}

// Check if string contains substring (not case sensitive)
function checkForSubstring(fullStr, subStr) {
    fullStr = fullStr.toLowerCase();
    subStr = subStr.toLowerCase();
    if (fullStr.indexOf(subStr) !== -1) {
        return true;
    } else {
        return false;
    }
}

// Remove attribute
function removeAttr(selectors, attr) {
    for (var i = 0; i < selectors.length; i++) {
        selectors[i].removeAttribute(attr);
    }
}

// Check for class
function elHasClass(element, className) {
    return (' ' + element.className + ' ').indexOf(' ' + className + ' ') > -1;
}
