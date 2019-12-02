let Excel = require('exceljs');

let methods = {
    readExcel: function (file_name1, file_name2) {
// read from a file
        var file1 = [];
        let workbook = new Excel.Workbook();
        workbook.xlsx.readFile(file_name1);
        var file2 = [];
        var file3 = [];
        let filebools = [];
        let workbook2 = new Excel.Workbook();
        workbook2.xlsx.readFile(file_name2)
            .then(function () {
                let worksheet = workbook.getWorksheet("CS");
                worksheet.eachRow((row, rowNumber) => {
                    var course_data = new Object();
                    var attra = [];
                    let column_keys = worksheet.getRow(1).values;
                    column_keys.forEach((single_Name) => {
                        attra.push(single_Name)

                    });
                    column_keys.forEach((single_Name) => {
                        attra.push(single_Name + "_bool_StrikeThru")
                    });
                    var res = JSON.stringify(row.values).split(",");
                    var strucken = [];

                    for (var z = 1; z < res.length; z++) {
                        if (row.getCell(z).value !== null && row.getCell(z).font !== null && row.getCell(z).font.strike !== null) {

                            if (row.getCell(z).font.strike) {
                                //console.log("struck")
                                strucken.push(1);
                            } else {
                                //console.log("miss")
                                strucken.push(0);

                            }
                        } else {
                            //console.log("miss")
                            strucken.push(0);
                        }

                    }

                    var boolsize = column_keys.length;
                    var increment = 1;
                    var sturckenSpot = 0;
                    attra.forEach((spot) => {
                        if (increment >= boolsize) {
                            if (strucken[sturckenSpot] === 1) {
                                course_data[spot] = true;
                                sturckenSpot++;
                            } else {
                                course_data[spot] = false;
                                sturckenSpot++;
                            }

                        } else {
                            course_data[spot] = res[increment];
                            increment++;
                        }

                    });
                    file1.push(course_data);
                });

                let worksheet2 = workbook2.getWorksheet("CS");
                worksheet2.eachRow((row, rowNumber) => {
                    var course_data2 = new Object();
                    var attra2 = [];
                    var struken2 = [];
                    let column_keys2 = worksheet2.getRow(1).values;
                    column_keys2.forEach((single_Name) => {
                        attra2.push(single_Name)
                    });
                    column_keys2.forEach((single_Name) => {
                        attra2.push(single_Name + "_bool_StrikeThru")
                    });

                    for (var z = 1; z < column_keys2.length; z++) {
                        if (row.getCell(z).value !== null && row.getCell(z).font !== null && row.getCell(z).font.strike !== null) {

                            if (row.getCell(z).font.strike) {
                                //console.log("struck")
                                struken2.push(1);
                            } else {
                                //console.log("miss")
                                struken2.push(0);

                            }
                        } else {
                            //console.log("miss")
                            struken2.push(0);
                        }

                    }
                    var boolsize = column_keys2.length;
                    var res2 = JSON.stringify(row.values).split(",");
                    var increment2 = 1;
                    var sturckenSpot2 = 0;

                    attra2.forEach((spot) => {
                        if (increment2 >= boolsize) {
                            if (struken2[sturckenSpot2] === 1) {
                                course_data2[spot] = true;
                            } else {
                                course_data2[spot] = false;
                                sturckenSpot2++
                            }

                        } else {
                            course_data2[spot] = res2[increment2];
                            increment2++;
                        }
                    });
                    file2.push(course_data2);
                });
                keys1 = [];
                values1 = [];
                keys2 = [];
                values2 = [];
                var file3Red = [];

                //**Need to output the diffrences.
                for (var i = 0; i < file1.length; i++) {
                    keys1 = Object.keys(file1[i]);
                    keys2 = Object.keys(file2[i]);
                    values1 = Object.values(file1[i]);
                    values2 = Object.values(file2[i]);

                    redChecker = [];

                    for (var x = 0; x < values1.length; x++) {
                        if (values1[x] !== values2[x]) {
                            redChecker.push(true);
                            //console.log("Dif Found!")
                        } else {
                            redChecker.push(false);

                        }

                    }

                    var course_data3 = new Object();
                    var red_course_data = new Object();

                    var increment2 = 0;
                    var redIncrement = 0;

                    keys2.forEach((spot) => {

                        course_data3[spot] = values2[increment2];
                        increment2++;
                    });
                    keys2.forEach((spot) => {

                        red_course_data[spot] = redChecker[redIncrement];
                        redIncrement++;
                    });

                    file3.push(course_data3);
                    file3Red.push(red_course_data);
                    //console.log(redChecker)

                }
                //console.log(file3);
                console.log(file3Red);
                filebools = file3Red;

            });
        return {
            array1: file3,
            red_array: filebools
        };
    }
}

module.exports = methods;
