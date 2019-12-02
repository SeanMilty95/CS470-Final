let Excel = require('exceljs');
let Moment = require('moment'); //Using this to format time
const fs = require('fs');

//const testJSON = require('./test.json');
//const csDepartment = require('./CS_DepartmentView_2187_JSON');
//const csFullJSON = require('./CS_Courses_2197_Schd_Dept_Stu_Views');

const flatCourses = require('./courses_flat_pp');
const rank = require('./rank');

let methods = {

    writeToExcel: function()
    {
    // construct a streaming XLSX workbook writer with styles and shared strings
    let options = {
        filename: './streamed-workbook.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    let worksheet = workbook.addWorksheet('Test Sheet');
    worksheet.columns = [
        {header: 'Merged', key: 'm'},
        {header: 'Last', key: 'l'},
        {header: 'First', key: 'f'},
        {header: 'Age', key: 'age'}
    ];

    testJSON.forEach((object, i) => {
        worksheet.getRow(i + 2).values = ({l: object.l, f: object.f, age: object.age});
    });

    worksheet.mergeCells('A2:A4');
    worksheet.getCell('A2').value = 'This is a merged cell';

    /*worksheet.addRow({l: 'Smith', f: 'John', age: 26}).commit();
    worksheet.addRow({l: 'Smith', f: 'Fox', age: 35}).commit();
    worksheet.addRow({l: 'Schmidt', f: 'Michael', age: 24}).commit();*/

    worksheet.commit();
    // Finished the workbook.
    workbook.commit()
        .then(function () {
            // the stream has been written
        });

},

writeWithJson: function() {
    let colunms = ["subject", "catalog", "course_title", "units", "ftes"];
    let options = {
        filename: './json-workbook.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    let worksheet = workbook.addWorksheet('Test Sheet');

    //columns only work on the top most row of the excel file
    worksheet.columns = [
        {header: 'Subject', key: 'subject', style: {font: {name: 'Times New Roman'}}, width: 8},
        {header: 'Catalog', key: 'catalog', style: {font: {name: 'Times New Roman'}}},
        {header: 'Course Title', key: 'course_title', style: {font: {name: 'Times New Roman'}}, width: 35},
        {header: 'Units', key: 'units', style: {font: {name: 'Times New Roman'}}},
        {header: 'FTES', key: 'ftes', style: {font: {name: 'Times New Roman'}}}
    ];

    worksheet.getRow(1).font = {bold: true};

    let departmentKeys = Object.keys(csDepartment);
    //console.log(departmentKeys);
    let count = 0;
    departmentKeys.forEach(key => {
        //console.log(csDepartment[key][0]);
        count++;
        worksheet.addRow({
            subject: csDepartment[`${key}`][0].subject, catalog: csDepartment[`${key}`][0].catalog,
            course_title: csDepartment[`${key}`][0].course_title, units: csDepartment[`${key}`][0].units,
            ftes: csDepartment[`${key}`][0].ftes
        });
    });
    //=SUM(E2:E39)
    worksheet.addRow({units: "Total"});
    //This is how you had functions to excel
    worksheet.getCell(`E${count + 2}`).value = {formula: `=SUM(E2:E${count + 1})`};

    worksheet.commit();
    // Finished the workbook.
    workbook.commit()
        .then(function () {
            // the stream has been written
        });


},

produceColumns: function(orderRanks, columnTitles, columns_width) {
    //This function uses the keys of each object and converts them to a more presentable title for each column
    orderRanks.forEach(key => {
        let mod_key = key[0].toUpperCase() + key.slice(1);
        let i = mod_key.indexOf('_');
        while (i !== -1) {
            mod_key = mod_key.slice(0, i) + ' ' + mod_key[i + 1].toUpperCase() + mod_key.slice(i + 2);
            i = mod_key.indexOf('_');
        }
        columnTitles.push(mod_key);
        columns_width.push(mod_key.length); //Used to help make sure the column is large enough to view the name
    });
},

writeRows: function(orderRanks, flattenedObj, worksheet, columns_width, startRow) {
    //This function will write all the rows with no formatting, expect for start and end time
    let count = startRow;
    flattenedObj.forEach((obj) => {
        let Row_values = [];
        orderRanks.forEach((key, i) => {
            let value = obj[key];
            //8 and 9 correspond to start and end time in the rank.json
            if ((i === 8 || i === 9) && value != null) {
                //creates a string with the correct time format
                let formattedTime = new Moment(value, "HH:mm:ss").format("HH:mm A");
                console.log(formattedTime);
                Row_values.push(formattedTime);
            } else
                Row_values.push(value);
            //This is to make sure the we know how big to make the columns in order to fit the largest word
            if (obj[key] !== null && columns_width[i] < obj[key].length)
                columns_width[i] = obj[key].length;
        });
        console.log(Row_values);
        worksheet.getRow(count++).values = Row_values;
    });
    return count; //return so we know what row we stopped on
},

basicWriteCourse: function() {
    let options = {
        filename: './courses-workbook.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    let worksheet = workbook.addWorksheet('Basic Courses Sheet');

    const beginningRow = 2;
    let startRow = beginningRow;

    let columnTitles = [];
    let columns_width = [];

    methods.produceColumns(rank, columnTitles, columns_width);

    //console.log(columnTitles);

    worksheet.getRow(startRow++).values = (columnTitles);

    startRow = methods.writeRows(rank, flatCourses, worksheet, columns_width, startRow);

    //Adjust how widths of the columns are calculated
    let columns_properties = [];
    columns_width.forEach((value, i) => {
        columns_properties.push({width: value + 3, style: {font: {name: 'Times New Roman'}}});
    });
    worksheet.columns = columns_properties;

    //Modify the columns to make them centered, middle, and bold.
    let row = worksheet.getRow(beginningRow);
    row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    row.font = {bold: true};
    row.height = 30;

    //Add the Title
    let letter = String.fromCharCode(96 + columns_width.length).toUpperCase();
    worksheet.mergeCells(`A1:${letter}1`);
    worksheet.getCell('A1').value = "Computer Science Courses -- Fall 2019";
    let row2 = worksheet.getRow(1);
    row2.height = 45;
    row2.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    row2.font = {size: 15, bold: true};


    worksheet.commit();
    // Finished the workbook.
    workbook.commit()
        .then(function () {
            // the stream has been written
        });
},

writeCoursesByInstructor: function() {
    let intructor_keys = ["instructor_lName", "instructor_fName", "instructor_id"];
    let column_keys = ["subject", "catalog", "section", "component", "course_title", "instructor",
        "wtu", "units", "meeting_pattern", "start_time", "end_time", "facility_name", "total_enrolled", "ftes"];

    let options = {
        filename: './instructor-workbook.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    let worksheet = workbook.addWorksheet('Instructor Courses Sheet');


    let columnTitles = [];
    let columns_width = [];

    methods.produceColumns(column_keys, columnTitles, columns_width);

    //get only unique ids for instructors
    let courses_instructors = {};
    flatCourses.forEach((obj) => {
        if (courses_instructors[obj.instructor_lName] === undefined)
            courses_instructors[obj.instructor_lName] = [];
        courses_instructors[obj.instructor_lName].push(obj);
    });

    let row = 2;
    let list_instructors = Object.keys(courses_instructors).sort();
    list_instructors.forEach(lname => {
        let full_name = courses_instructors[lname][0].instructor_lName + ', ' + courses_instructors[lname][0].instructor_fName;
        let id = '(' + courses_instructors[lname][0].instructor_id + ')';

        let letter = String.fromCharCode(96 + column_keys.length).toUpperCase();
        worksheet.mergeCells(`A${row}:${letter}${row}`);
        worksheet.getRow(row).height = 35;

        let instructor_cell = worksheet.getCell(`A${row++}`);
        instructor_cell.value = (full_name + id);
        instructor_cell.font = {name: 'Times New Roman', size: 14, bold: true};
        instructor_cell.alignment = {vertical: 'middle'};
        instructor_cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            bgColor: {argb: 'FF696969'}
        };

        let ids_row = worksheet.getRow(row++);
        ids_row.values = columnTitles;
        ids_row.height = 25;
        ids_row.font = {name: 'Times New Roman', bold: true};
        ids_row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};

        courses_instructors[lname].forEach(obj => {
            let Row_values = [];
            column_keys.forEach((key, i) => {
                if (key === 'instructor') {
                    Row_values.push(full_name);
                    if (columns_width[i] < full_name.length)
                        columns_width[i] = full_name.length;
                } else if ((i === 9 || i === 10) && obj[key] != null) {
                    //creates a string with the correct time format
                    let formattedTime = new Moment(obj[key], "HH:mm:ss").format("HH:mm A");
                    //console.log(formattedTime);
                    Row_values.push(formattedTime);
                } else {
                    Row_values.push(obj[key]);
                    if (obj[key] !== null && columns_width[i] < obj[key].length)
                        columns_width[i] = obj[key].length;
                }
            });
            //console.log(Row_values);
            worksheet.getRow(row++).values = Row_values;
        });
        row++ //skip a row
    });

    //Adjust how widths of the columns are calculated
    let columns_properties = [];
    columns_width.forEach((value, i) => {
        columns_properties.push({width: value + 3, style: {font: {name: 'Times New Roman'}}});
    });
    worksheet.columns = columns_properties;

    worksheet.commit();
    // Finished the workbook.
    workbook.commit()
        .then(function () {
            // the stream has been written
        });

},


flattenScheduleJSON: function() {
    const scheduleJSON = csFullJSON.schedulerView;
    const base_class_keys = Object.keys(scheduleJSON);


    let flattenedArray = [];
    base_class_keys.forEach(bc_key => {
        let class_keys = Object.keys(scheduleJSON[bc_key]);
        class_keys.forEach(c_key => {
            if (c_key === 'isMultiComponent')
                console.log("Do Something Here");
            else {
                scheduleJSON[bc_key][c_key].forEach(obj => {
                    let flattenedObject = {};
                    let tempArray_Components = [];
                    let object_keys = Object.keys(obj);
                    object_keys.forEach(o_key => {
                        //get components for this class
                        let components = obj.components;
                        if (components.includes(o_key)) {
                            let flatComp = {};
                            let comp_keys = Object.keys(obj[o_key]);
                            comp_keys.forEach(comp_key => {
                                //we can probably put these two if statements together
                                if (comp_key === "instructors") {
                                    obj[o_key][comp_key].forEach(instructor_obj => {
                                        let instructor_keys = Object.keys(instructor_obj);
                                        //Check if instructor was changed
                                        instructor_keys.forEach(i_key => {
                                            flatComp[i_key] = instructor_obj[i_key];
                                        })
                                        //console.log(instructor_keys);
                                    })
                                } else if (comp_key === "meeting_pattern") {
                                    obj[o_key][comp_key].forEach(meeting_obj => {
                                        let meeting_keys = Object.keys(meeting_obj);
                                        //Check if meeting times were changed
                                        meeting_keys.forEach(m_key => {
                                            flatComp[m_key] = meeting_obj[m_key];
                                        })
                                        //console.log(meeting_keys);
                                    })

                                } else {
                                    flatComp[comp_key] = obj[o_key][comp_key];
                                }
                            });
                            flatComp["component"] = o_key;
                            tempArray_Components.push(flatComp);
                        } else if (o_key !== "components") {
                            flattenedObject[o_key] = obj[o_key];
                        }
                    });
                    tempArray_Components.forEach(component => {
                        //Shallow copy, only works because object is fully flat
                        let copy = Object.assign({}, flattenedObject);
                        let flat = Object.assign(copy, component);
                        flattenedArray.push(flat);
                    });
                })
            }
        });
    });

    //Sort by catalog and section
    flattenedArray.sort((a, b) => (a.catalog > b.catalog) ? 1 : (a.catalog === b.catalog) ? ((a.section > b.section) ? 1 : -1) : -1);

    //fs.writeFile("courses_flat_created.json", JSON.stringify(flattenedArray, null, 2));

},


excelToJSON: function(filename, flatJSONArray) {
    /*
    * This function take the name of the excel file the you want to read in(must be in the same directory)
    * and opens that excel sheet. It then reads from it all the information needed to generated a weekly
    * schedule. NOTE this function's excel must be formatted with the exact same column ids/names as seen
    * below or it will not work.
    * It returns an array of flat JSON objects
    */

    //HARD CODED
    let column_name = ["Subject", "Catalog", "Component", "Section", "Last", "Facil ID", "START TIME", "END TIME", "Pat", "Auto Enrol"];
    let column_keys = {
        "Subject": "subject", "Catalog": "catalog", "Component": "component", "Section": "section",
        "Last": "instructor_lName", "Facil ID": "facility_name", "START TIME": "start_time", "END TIME": "end_time",
        "Pat": "meeting_pattern", "Auto Enrol": "auto_enroll"
    };

    let workbook = new Excel.Workbook();
    return workbook.xlsx.readFile(filename)
        .then(function () {
            //we are assuming there is only one worksheet, other access attempts didn't work
            workbook.eachSheet((worksheet, sheetId) => {
                let color_index = 0;
                let time = null;
                let lookup_table = {};

                let column_ids = worksheet.getRow(1).values;
                worksheet.eachRow((row, index) => {
                    if (index !== 1) {
                        let JSON = {};
                        column_ids.forEach((id, index) => {
                            if (column_name.includes(id)) {
                                let cell = row.getCell(index);
                                if (JSON[column_keys[id]] === undefined) {
                                    if (id === "START TIME" || id === "END TIME")
                                        JSON[column_keys[id]] = new Moment(cell.value, "hh:mm A").format("HH:mm:ss");
                                    else
                                        JSON[column_keys[id]] = cell.value;
                                }
                            }
                        });

                        //This add a number to the classes that can be used to connect classes by color
                        //MUST be ordered by sections
                        if (JSON["auto_enroll"] !== null) {
                            let new_time = JSON["catalog"] + '_' + JSON["start_time"] + '_' + JSON["end_time"] + '_' + JSON["meeting_pattern"];
                            if (time !== new_time) {
                                color_index++;
                                time = new_time
                            }
                            lookup_table[JSON["catalog"] + '_' + JSON["auto_enroll"]] = color_index;
                            JSON.color_class_idx = color_index;

                        } else if (lookup_table[JSON["catalog"] + '_' + JSON["section"]] !== undefined) {
                            JSON.color_class_idx = lookup_table[JSON["catalog"] + '_' + JSON["section"]];
                        } else {
                            let new_time = JSON["catalog"] + '_' + JSON["start_time"] + '_' + JSON["end_time"] + '_' + JSON["meeting_pattern"];
                            if (time !== new_time) {
                                color_index++;
                                time = new_time
                            }
                            JSON.color_class_idx = color_index;
                        }

                        flatJSONArray.push(JSON)
                    }
                });
            });
        });
},

reduceJSON: async function(filename, reducedJSON) {
    /*
    * This function takes a file name of a JSON in the current directory (no ./ is need in the name) and
    * reads from it. The file must be an array of flat JSON objects that holds information needed to generate
    * a weekly schedule. NOTE auto_enrol does not exist in JSON form yet.
    * This will return another array of flat objects that has only the info needed for the weekly schedule
    */

    //Currently has no way to add color by class, needs auto_enroll in the JSON
    let name_keys = ["subject", "catalog", "component", "section", "instructor_lName", "facility_name", "meeting_pattern",
        "start_time", "end_time"];

    let fullJSON = require("./" + filename);
    fullJSON.forEach(obj => {
        if (obj.meeting_pattern !== "ARR") {
            let reduced_item = {};
            name_keys.forEach(name_key => {
                reduced_item[name_key] = obj[name_key]
            });
            reducedJSON.push(reduced_item);
        }
    });
},

patternToMeetingArray: function(pattern) {
    /*
    * SUPPORT FUNCTION FOR: organizeWeeklySchedule
    *
    * This function is passed a string of letters that represent days of the week, in the form of
    * "MTWTHF". This is need only because TH is two letters, so a straight up split wouldn't work.
    * Formats and returns an array so meeting pattern is in the form of [ 'M', 'T', 'W', 'TH', 'F' ]
    */
    let patternArray = pattern.split('');
    patternArray.forEach((char, index) => {
        if (char === 'H') {
            patternArray[index - 1] = 'TH';
            patternArray.splice(index, 1);
        }
    });
    return patternArray;
},

reviewConflictingItems: function(existing_item, schedule_item) {
    /*
    * SUPPORT FUNCTION FOR: organizeWeeklySchedule
    *
    * This function takes two objects and makes comparisons on the start_time and end_time of both the objects
    * and determines if their is a conflict between the existing_time and the item to be added(schedule_item).
    * If the times are the exact same, it then checks if the class is just another section of the existing class
    * and it makes a note just in case it is needed latter.
    * This will return a number, 0,1, or 2. 0 means no conflict, 1 means same class, another section, and
    * 2 means conflict between the two different classes
    */

    let name_keys = ["subject", "catalog", "component", "section", "instructor_lName", "facility_name"];
    let possible_to_push = 0;

    let E_s_time = existing_item["start_time"];
    let E_e_time = existing_item["end_time"];

    let S_s_time = schedule_item["start_time"];
    let S_e_time = schedule_item["end_time"];

    if (E_s_time === S_s_time && E_e_time === S_e_time) {
        possible_to_push = 1;
        name_keys.forEach(name_key => {
            if (name_key !== "section" && existing_item[name_key] !== schedule_item[name_key])
                possible_to_push = 2;
        });
        if (possible_to_push === 1)
            existing_item["multi_section"] = true;
    } else if (E_s_time <= S_s_time && E_e_time >= S_s_time) {
        //console.log("START time conflict", s_time, e_time, schedule_item["start_time"]);
        possible_to_push = 2;
    } else if (E_s_time <= S_e_time && E_e_time >= S_e_time) {
        //console.log("END time conflict", s_time, e_time, schedule_item["end_time"]);
        possible_to_push = 2;
    }

    return possible_to_push;
},

organizeWeeklySchedule: function(flatJSONArray) {
    /*
    *
    */

    let meeting_days = {M: "monday", T: "tuesday", W: "wednesday", TH: "thursday", F: "friday"};
    let weekly = {monday: [[]], tuesday: [[]], wednesday: [[]], thursday: [[]], friday: [[]]};

    let instructor_index = 0;
    let instructor_to_index = {};

    //flatJSONArray is the simple Array of flat JSONs
    flatJSONArray.forEach(obj => {
        let schedule_item = obj;

        //color is added by instructor
        if (instructor_to_index[obj["instructor_lName"]] === undefined) {
            instructor_to_index[obj["instructor_lName"]] = instructor_index;
            instructor_index++;
        }

        //move to a different function or remove
        //name_keys.forEach( name_key => { schedule_item[name_key] = obj[name_key] });
        //time_keys.forEach( time_key => { schedule_item[time_key] = obj[time_key] });

        //color by instructor
        schedule_item.color_instructor_idx = instructor_to_index[obj["instructor_lName"]];

        let meeting_array = methods.patternToMeetingArray(obj.meeting_pattern);
        meeting_array.forEach(day_letter => {
            let day_full = meeting_days[day_letter];
            if (day_full !== undefined) {
                let item_pushed = false;
                let index = 0;
                while (index < weekly[day_full].length) {
                    let column = weekly[day_full][index];
                    if (!item_pushed) {
                        if (column.length === 0) {
                            column.push(schedule_item);
                            item_pushed = true;
                        } else {
                            let possible_to_push = 0;
                            column.forEach(existing_item => {
                                if (possible_to_push === 0) {
                                    possible_to_push = methods.reviewConflictingItems(existing_item, schedule_item);
                                }
                            });
                            if (possible_to_push === 0) {
                                column.push(schedule_item);
                                item_pushed = true;
                            } else if (index + 1 === weekly[day_full].length && possible_to_push === 2) {
                                weekly[day_full].push([]);
                            }
                        }
                    }
                    index++;
                }
            }
        });
    });
    //return values are the organized weekly schedule and an index for how many colors need to be generated
    return [weekly, instructor_index];
},

militaryTime: function() {
    //Just sets up a basic array of timestamps for the excel schedule sheet
    let startTime = new Moment('08:00:00', "HH:mm:ss");
    let timeArray = [];
    while (startTime.hours() !== 22) {
        timeArray.push(startTime.clone());
        startTime.add(5, "m")
    }
    return timeArray;
},

generateColors: function(size) {
    let colors = [];
    for (let i = 0; i < size; i++) {
        colors.push(Math.random().toString(16).substr(-6))
    }
    colors.sort();
    return colors;
},

generateColumnLetters: function(size) {
    let integer_letter = 65; // A
    let integer_letter2 = 65; // A
    let letters = [];
    let doubleLetter = false;
    for (let i = 0; i < size; i++) {
        if (doubleLetter) {
            letters.push(String.fromCharCode(integer_letter2) + String.fromCharCode(integer_letter));
        } else
            letters.push(String.fromCharCode(integer_letter));
        integer_letter++;
        if (integer_letter >= 91) {
            integer_letter = 65;
            doubleLetter = true;
        }
    }
    return letters;
}

,

writeWeeklySchedule: function(organizeWeekly, color_type) {
    let options = {
        filename: './weekly-schedule.xlsx',
        useStyles: true,
        useSharedStrings: true
    };
    let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    let worksheet = workbook.addWorksheet('Weekly Layout');

    let name_keys = ["subject", "catalog", "component", "section", "instructor_lName", "facility_name"];

    let column_values = ["TIMES"];

    let values = organizeWeekly;
    let weekly_schedule = values[0];
    let weekly_keys = ["monday", "tuesday", "wednesday", "thursday", "friday"];

    let colors = null;
    if (color_type === "instructor")
        colors = methods.generateColors(values[1]);
    else
        colors = methods.generateColors(35);

    let times = methods.militaryTime();
    let rows_for_times = {};
    let row_index = 2;

    times.forEach(time => {
        let formatted_time = time.format("HH:mm:ss");
        rows_for_times[formatted_time] = row_index;

        if (time.minutes() % 15 === 0) {
            let column_index = 1;
            let rowValues = [];
            let column_time = time.format("hh:mm A");
            rowValues[column_index] = column_time;
            weekly_keys.forEach(day => {
                column_index += weekly_schedule[day].length + 3;
                rowValues[column_index] = column_time;
            });
            worksheet.getRow(row_index).values = rowValues;
            row_index++;
        }
    });
    //letters should hold enough column letters as not to error out
    let letters = methods.generateColumnLetters(52);
    let letter_index = 2; // C
    weekly_keys.forEach(day => {
        let starting_letter = letters[letter_index];
        worksheet.getColumn(letter_index).width = 1;
        column_values[letter_index] = day.toUpperCase();
        weekly_schedule[day].forEach((column, index) => {
            let letter = letters[letter_index];
            letter_index++;
            worksheet.getColumn(letter).width = 15;
            column.forEach(obj => {
                let start_row = rows_for_times[obj.start_time];
                let end_row = rows_for_times[obj.end_time];
                let color = null;
                if (color_type === "instructor")
                    color = colors[obj.color_instructor_idx];
                else
                    color = colors[obj.color_class_idx];

                worksheet.mergeCells(`${letter}${start_row}`, `${letter}${end_row - 1}`);
                let cell = worksheet.getCell(`${letter}${start_row}`);

                let cell_text = "";
                name_keys.forEach(key => {
                    cell_text += obj[key] + " "
                });

                cell.border = {
                    top: {style: 'medium'},
                    left: {style: 'medium'},
                    bottom: {style: 'medium'},
                    right: {style: 'medium'}
                };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {argb: color}
                };
                cell.value = cell_text;
                cell.font = {name: 'Times New Roman', bold: true};
                cell.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
            });
        });
        worksheet.getColumn(letter_index + 1).width = 1;
        column_values[letter_index + 1] = "TIME";
        worksheet.mergeCells(`${starting_letter}1`, `${letters[letter_index - 1]}1`);
        letter_index += 3
    });

    //only does monday
    /*weekly_schedule.friday.forEach( (column, index) => {
        //start letter at C
        let letter = String.fromCharCode(65 + index + 2);
        worksheet.getColumn(letter).width = 15;
        column.forEach(obj => {
            let start_row = rows_for_times[obj.start_time];
            let end_row = rows_for_times[obj.end_time];
            worksheet.mergeCells(`${letter}${start_row}`,`${letter}${end_row-1}`);
            let cell = worksheet.getCell(`${letter}${start_row}`);
            let cell_text = "";
            Object.keys(obj).forEach(key => {
                cell_text += obj[key] + " "
            });
            let color = colors[obj.color_idx];
            cell.border ={
                top: {style:'medium'},
                left: {style:'medium'},
                bottom: {style:'medium'},
                right: {style:'medium'}
            };
            cell.fill = {
                type: 'pattern',
                pattern:'solid',
                fgColor: {argb:color}
            };
            cell.value = cell_text;
            cell.font = {name: 'Times New Roman', bold: true};
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        });
    });*/
    let row = worksheet.getRow(1);
    row.values = column_values;
    row.font = {name: 'Times New Roman', bold: true};
    row.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
    row.border = {bottom: {style: 'medium'}};

    worksheet.commit();
    // Finished the workbook.
    workbook.commit()
        .then(function () {
            // the stream has been written
        });
},

readAndExecute: async function(filename, color_type) {
    let myJSONArray = [];
    let foundString = filename.substring(filename.lastIndexOf('.') + 1);
    if (foundString === "xlsx") {
        //When using excel the time is in a different format
        await methods.excelToJSON(filename, myJSONArray).catch(err => console.log(err));
    } else if (foundString === "json")
        await methods.reduceJSON(filename, myJSONArray).catch(err => console.log(err));
    else {
        console.log(filename + " is a file that doesn't end with json or xlsx")
    }

    if (myJSONArray.length !== 0) {
        let values = methods.organizeWeeklySchedule(myJSONArray);
        //console.log(values[0]);
        methods.writeWeeklySchedule(values, color_type);
    }
}

};
module.exports = methods;
//readAndExecute("courses_flat_created.json", "instructor").catch(err => console.log(err));