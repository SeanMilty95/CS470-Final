const Diff = require('./excel_manip.js');
let Week = require('./WeekView.js');
let all_dif = true; // Used to check if all diffs should be recorded or just the diffs of the editable attributes
let week_view = false; // Used to check if the week view should be generated

function main() {
    if (process.argv.length === 5){
        //Do things based on the flags set
        if(process.argv[2].toUpperCase() === 'E'){
            all_dif = false;
            Generate_Dif(process.argv[3], process.argv[4]);
        }
        else if (process.argv[2].toUpperCase() === 'W'){
            week_view = true;
            Generate_Dif(process.argv[3], process.argv[4]);
        }
        else if (process.argv[2].toUpperCase() === 'D'){
            Generate_Dif(process.argv[3], process.argv[4]);
        }
        else{
            console.log('The option entered is not valid.\nChoose either D, E, or W');
        }
    }
    else if( process.argv.length === 4){
        //Do things normally with default values
        Generate_Dif(process.argv[2], process.argv[3]);
    }
    else{
        console.log('An incorrect number of arguments have been entered!');
        console.log('Please use the form: \nindex.js {option} {file1.xlsx} {file2.xlsx}');
    }
}

function Generate_Dif(file1, file2){
    if(week_view){
        //Run Michaels program
        new Week.readAndExecute(file1, "Instructors");
        console.log('Week view generating...');
    }
    else {
        console.log('File name 1 -> ' + file1);
        console.log('File name 2 -> ' + file2);
        let diffObject = new Diff(file1, file2, all_dif);
        diffObject.Read_and_generate();
        //diffObject.Compare_Worksheets();
        //Always call this function as it commits all workbooks associated with the difference object
        diffObject.Commit_All();
    }

}


main();