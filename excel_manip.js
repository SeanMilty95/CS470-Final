let Excel = require('exceljs');
let reader = require('./excelTesting.js');

class DiffClass {
    constructor(file_name1, file_name2, all){
        //this.read_workbook1 = this.Make_Workbook_Stream(file_name1);
        //this.read_workbook2 = this.Make_Workbook_Stream(file_name2);
        //this.diff_workbook = this.Make_Workbook_Stream('C:\\Users\\milte\\Desktop\\Excel-Output.xlsx');
        //this.read_sheet1 = this.read_workbook1.getWorksheet('CS');
        //this.read_sheet2 = this.read_workbook2.getWorksheet('CS');
        //this.diff_sheet = this.diff_workbook.addWorksheet('CS')

        this.all_diffs = all;
        this.diff_array = [];
        this.file1 = file_name1;
        this.file2 = file_name2;
        console.log('Class constructed...');
    };
    Make_Workbook_Stream(file_name){
        //Creates a workbook stream from the filename passed to the function.
        let options = {
            filename: file_name,
            useStyles: true,
            useSharedStrings: true
        };
        return new Excel.stream.xlsx.WorkbookWriter(options);
        //return workbook;
    };
    Read_and_generate(){
       let arrays = reader.readExcel(this.file1, this.file2); //This function returns an object of 2 arrays that contain diff values
       let first = arrays.array1;
       this.diff_array = arrays.red_array;

       //Uncomment this when ready to create a file of differences
       //this.Create_Diff_File();

    }
    Compare_Worksheets(){
        /*
        console.log(this.read_sheet1);
        let row1 = this.read_sheet1.getRow(2);
        let row2 = this.read_sheet2.getRow(2);
        if(row1 !== row2){
            console.log("Rows are not equal!");
        }
        else{
            console.log("Rows are equal or cannot be compared.\n Testing...\n");
            console.log(row1);
            console.log(row2);
        }
        */
    };
    Create_Diff_File(){
        let options = {
            filename: 'C:\\Users\\milte\\Desktop\\cs470Output.xlsx',//Change the filepath to either generic desktop or
            //user given filepath.
            useStyles: true,
            useSharedStrings: true
        };
        let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

        let sheet1 = workbook.addWorksheet('CS');
        sheet1.columns = [
            {header: 'Term', key: 'term', width: 15},
            {header: 'Acad Group', key: 'group', width: 15},
            {header: 'Subject', key: 'subj', width: 15},
            {header: 'Catalog', key: 'catalog', width: 15},
            {header: 'Class Nbr', key: 'class_nbr', width: 15},
            {header: 'Section', key: 'section', width: 15},
            {header: 'Min Units', key: 'min_units', width: 15},
            {header: 'Designation', key: 'designation', width: 15},
            {header: 'Mtg Start', key: 'mtg_start', width: 15},
            {header: 'Mtg End', key: 'mtg_end', width: 15},
            {header: 'Pat', key: 'pat', width: 15},
            {header: 'Facil ID', key: 'facil_id', width: 15},
            {header: 'Cap Enrl', key: 'cap_enrl', width: 15},
            {header: 'Tot Enrl', key: 'tot_enrl', width: 15},
            {header: 'Class Type', key: 'class_type', width: 15}
        ];

        sheet1.getRow(1).fill = {// Makes the first row grey
            type: "pattern",
            pattern: "solid",
            fgColor: {argb: '808080'},
            bgColor: {argb: '808080'}
        };
        sheet1.getRow(1).border ={
            top: {style: "medium", color: {argb: '000000'}},
            bottom: {style: "medium", color: {argb: '000000'}},
            left: {style: "medium", color: {argb: '000000'}},
            right: {style: "medium", color: {argb: '000000'}}
        };

        for (let i = 0; i < this.diff_array.length; i++){
            sheet1.addRow(this.diff_array[i]);
        }
        sheet1.eachRow(function(row, rowNumber){
            row.eachCell(function(cell, colNumber){
                cell.alignment = {horizontal: "left"};
            });
        });

        sheet1.commit();
    }
    Commit_All(){
        //this.read_workbook1.commit();
        //this.read_workbook2.commit();
        //this.diff_workbook.commit();//This is now committed after its creation
        console.log("All commited");
    };

}

module.exports = DiffClass;