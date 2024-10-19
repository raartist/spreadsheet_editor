import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import ImportBtn from '../Import';

const ExcelEditor = () => {
    const [file, setFile] = useState(null);
    const [data, setData] = useState([]);
    const [editingCell, setEditingCell] = useState({ row: null, col: null });
    const [editingValue, setEditingValue] = useState('');
    const [draggedCell, setDraggedCell] = useState({ row: null, col: null });
    // const [draggedColIndex, setDraggedColIndex] = useState(null);

    const handleFileChange = (event) => {
        const uploadedFile = event.target.files[0];
        if (uploadedFile) {
            setFile(uploadedFile);
            readExcel(uploadedFile);
        }
    };

    const readExcel = async (file) => {
        const workbook = new ExcelJS.Workbook();
        const data = await workbook.xlsx.load(file);
        const worksheet = data.worksheets[0]; // Get the first worksheet
        const rows = [];

        worksheet.eachRow((row, rowNumber) => {
            let len = 6;
            let rowItems = [...row.values];
            if (rowItems.length < len) {
                for (let i = rowItems.length; i < len; i++) {
                    rowItems.splice(i, 0, "")
                }
            }
            rows.push(rowItems);
        });

        setData(rows);
    };

    const handleDownload = async () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Edited Data');

        data.forEach((row) => {
            worksheet.addRow(row);
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'edited.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    };
    const handleCellClick = (rowIndex, colIndex, value) => {
        setEditingCell({ row: rowIndex, col: colIndex });
        setEditingValue(value);
    };

    const handleChange = (event) => {
        setEditingValue(event.target.value);
    };

    const handleBlur = () => {
        const newData = [...data];
        newData[editingCell.row][editingCell.col] = editingValue;
        setData(newData);
        setEditingCell({ row: null, col: null });
    };

    const handleKeyEvents = (event, rowIndex, colIndex) => {
        if (event.key === 'Enter') {
            handleBlur();
        } else if (event.key === 'ArrowDown') {
            const nextRow = rowIndex + 1 < data.length ? rowIndex + 1 : rowIndex;
            setEditingCell({ row: nextRow, col: colIndex });
            setEditingValue(data[nextRow][colIndex] || '');
            event.preventDefault();
        } else if (event.key === 'ArrowUp') {
            const prevRow = rowIndex > 0 ? rowIndex - 1 : rowIndex;
            setEditingCell({ row: prevRow, col: colIndex });
            setEditingValue(data[prevRow][colIndex] || '');
            event.preventDefault();
        } else if (event.key === 'ArrowRight') {
            const nextCol = colIndex + 1 < data[rowIndex].length ? colIndex + 1 : colIndex;
            setEditingCell({ row: rowIndex, col: nextCol });
            setEditingValue(data[rowIndex][nextCol] || '');
            event.preventDefault();
        } else if (event.key === 'ArrowLeft') {
            const prevCol = colIndex > 0 ? colIndex - 1 : colIndex;
            setEditingCell({ row: rowIndex, col: prevCol });
            setEditingValue(data[rowIndex][prevCol] || '');
            event.preventDefault();
        }
    };
    //for cells
    const handleDragStart = (rowIndex, colIndex) => {
        setDraggedCell({ row: rowIndex, col: colIndex });
    };

    const handleDragOver = (event) => {
        event.preventDefault();
    };

    //for cells
    const handleDrop = (rowIndex, colIndex) => {
        const newData = [...data];
        const draggedValue = newData[draggedCell.row][draggedCell.col];

        newData[draggedCell.row][draggedCell.col] = newData[rowIndex][colIndex];
        newData[rowIndex][colIndex] = draggedValue;

        setData(newData);
        setDraggedCell({ row: null, col: null });
    };

    //for colums
    // const handleDragStart = (colIndex) => {
    //     setDraggedColIndex(colIndex);
    // };

    //for colums
    // const handleDrop = (targetColIndex) => {
    //     if (draggedColIndex === null || draggedColIndex === targetColIndex) return;
    //     const newData = [...data];
    //     const draggedColumn = newData.map(row => row[draggedColIndex]);
    //     newData.forEach((row, index) => {
    //         if (targetColIndex < draggedColIndex) {
    //             row.splice(targetColIndex, 0, row.splice(draggedColIndex, 1)[0]);
    //         } else {
    //             row.splice(targetColIndex, 0, row.splice(draggedColIndex, 1)[0]);
    //         }
    //     });

    //     setData(newData);
    //     setDraggedColIndex(null);
    // };

    return (
        <div>
            <ImportBtn handleFileChange={handleFileChange} />
            <button onClick={handleDownload} className="download-btn">Download Edited Excel</button>
            <table>
                <tbody>
                    {data.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                            {row.map((cell, colIndex) => (
                                <td key={colIndex}
                                    onClick={() => handleCellClick(rowIndex, colIndex, cell)}
                                    draggable
                                    onDragStart={() => handleDragStart(rowIndex, colIndex)}
                                    onDragOver={handleDragOver}
                                    onDrop={() => handleDrop(rowIndex, colIndex)}
                                >
                                    {editingCell.row === rowIndex && editingCell.col === colIndex ? (
                                        <input
                                            value={editingValue}
                                            onChange={handleChange}
                                            onBlur={handleBlur}
                                            onKeyDown={(event) => handleKeyEvents(event, rowIndex, colIndex)}
                                            autoFocus
                                        />
                                    ) : (
                                        cell
                                    )}
                                </td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

export default ExcelEditor;
