import React, {useState} from 'react';
import * as XLSX from 'xlsx';
import './App.css'
import {FaFileExcel, FaTimes} from 'react-icons/fa'; // Excel file icon and close icon from react-icons


type Weekdays = {
    monday: number;
    tuesday: number;
    wednesday: number;
    thursday: number;
    friday: number;
    saturday: number;
    sunday: number;
};

const App: React.FC = () => {
    const [file, setFile] = useState<File | null>(null);

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        const uploadedFile = event.target.files?.[0] || null;
        setFile(uploadedFile);
    };

    const uploadFile = () => {
        const fileInput = document.getElementById('fileID') as HTMLInputElement;
        if (fileInput) {
            fileInput.click(); // This triggers the file input to open
        }
    }

    const clearFile = () => {
        setFile(null);
        const fileInput = document.getElementById('fileID') as HTMLInputElement;
        if (fileInput) {
            fileInput.value = ''; // Clear the file input field
        }
    };

    const processExcel = () => {
        if (!file) {
            alert('Please upload an Excel file.');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e: ProgressEvent<FileReader>) => {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, {type: 'array'});

            const sheetA = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});

            const sheetB = mapSheet(sheetA);

            const newWorkbook = XLSX.utils.book_new();
            const newSheet = XLSX.utils.aoa_to_sheet(sheetB);
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet');

            XLSX.writeFile(newWorkbook, 'QCIF_format.xlsx');
        };

        reader.readAsArrayBuffer(file);
    };

    // Define types for sheet data processing
    const mapSheet = (sheetA: unknown[]): string[][] => {
        const sheetB: string[][] = [];
        sheetB.push(['Project Code', 'Project SubCode', 'Category', 'Task', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']);

        let projectCode;
        const resetWeekdays = () => ({
            monday: 0, tuesday: 0, wednesday: 0, thursday: 0, friday: 0, saturday: 0, sunday: 0
        });

        let weekdays: Weekdays = resetWeekdays();
        for (let i = 3; i < sheetA.length; i++) {
            const row: any = sheetA[i];
            if (row.length === 1) {
                projectCode = row[0] || '';
            } else if (row.length === 11) {
                const task = row[0];
                const category = row[3] ? row[3] : 'NA';
                const dayOfWeek = getDayOfWeek(row[4]);
                const timeSpent = row[9] || 0;

                weekdays[dayOfWeek.toLowerCase() as keyof Weekdays] = timeSpent;

                const existingIndex = sheetB.findIndex((entry) => entry[3] === task);
                if (existingIndex > -1) {
                    sheetB[existingIndex] = sheetB[existingIndex].map((value, index) => {
                        if (index >= 4 && index <= 10) {
                            return (
                                parseFloat(value || '0') +
                                (weekdays[getDayByIndex(index - 4).toLowerCase() as keyof Weekdays] || 0)
                            ).toString();
                        }
                        return value;
                    });
                    weekdays = resetWeekdays();
                } else {
                    sheetB.push([
                        projectCode,
                        '',
                        category,
                        task,
                        weekdays.monday.toString(),
                        weekdays.tuesday.toString(),
                        weekdays.wednesday.toString(),
                        weekdays.thursday.toString(),
                        weekdays.friday.toString(),
                        weekdays.saturday.toString(),
                        weekdays.sunday.toString(),
                    ]);
                    weekdays = resetWeekdays();
                }
            }

        }

        return sheetB;
    };


    // Map day index to weekday
    const getDayByIndex = (index: number) => {
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        return days[index];
    };

    const excelDateToJSDate = (serial: number) => {
        // Excel considers 1900 as the starting point, so we need to adjust.
        const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
        const millisecondsPerDay = 24 * 60 * 60 * 1000;

        // Calculate the number of milliseconds since Excel epoch and add it to the base date
        const dateInMs = excelEpoch.getTime() + serial * millisecondsPerDay;

        return new Date(dateInMs);
    }

    const getDayOfWeek = (serial: number) => {
        const jsDate = excelDateToJSDate(serial);

        // Array to map the day index (0 for Sunday, 1 for Monday, etc.)
        const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

        // Get the day of the week (0 = Sunday, 1 = Monday, etc.)
        const dayIndex = jsDate.getDay();

        // Return the corresponding day of the week
        return daysOfWeek[dayIndex];
    }

    return (
        // <div className="App">
        //     <h1>Excel Sheet Mapper</h1>
        //     <input type="file" onChange={handleFileUpload}/>
        //     <button onClick={processExcel}>Upload & Process</button>
        // </div>

        <>
            <div className="main-container">
                <div className="container">
                    <h1>Time Sheet Converter</h1>
                </div>
                <div className="container">
                    <div className="card">
                        <h3>Upload Files</h3>
                        <div className="drop_box">
                            {file ? <div>
                                <header>
                                    <h4>Selected File</h4>
                                </header>
                                <div className="file-preview">
                                    {file && (
                                        <div className="file-info">
                                            <FaFileExcel className="file-icon"/> {/* Excel file icon */}
                                            <span className="file-name">{file.name}</span>
                                            <FaTimes className="file-remove" onClick={clearFile}/> {/* Close button */}
                                        </div>
                                    )}
                                </div>
                            </div> : <div>
                                <header>
                                    <h4>Select File here</h4>
                                </header>
                                <p>Files Supported: XLSX, CSV</p>
                                <input type="file" hidden accept=".xlsx,.csv" id="fileID" onChange={handleFileUpload}/>
                                <button className="btn" onClick={uploadFile}>Choose File</button>
                            </div>
                            }
                        </div>
                    </div>
                </div>
                <div className="container">

                </div>
                <div className="container">
                    <button className="btn" onClick={processExcel}>Convert & Download</button>
                </div>
            </div>
        </>


    );
};

export default App;
