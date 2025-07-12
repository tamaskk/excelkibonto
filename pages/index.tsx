import { useState, useRef, useCallback, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface ExcelData {
  sheetName: string;
  data: any[][];
}

// Function to convert Excel serial date to JavaScript Date
const excelSerialToDate = (serial: number): Date => {
  // Days between Jan 1, 1900 and Jan 1, 1970
  const excelEpochOffset = 25569;
  // Adjust for Excel's 1900 leap year bug
  const leapYearBugAdjustment = 2;
  
  // Split serial into days and fractional time
  const days = Math.floor(serial) - (excelEpochOffset + leapYearBugAdjustment);
  const fraction = serial % 1;
  
  // Convert to milliseconds (days * 86400000 + fraction of day * 86400000)
  const milliseconds = (days + fraction) * 86400000;
  
  // Create JavaScript Date object
  return new Date(milliseconds);
};

const formatDecimal = (value: any): any => {
  if (value !== null && value !== undefined && !isNaN(Number(value))) {
    const numericValue = Number(value);
      const decimalPart = numericValue % 1;
      
      if (decimalPart < 0.5) {
        // Round down (floor)
        return Math.floor(numericValue);
      } else {
        // Round up (ceil)
        return Math.ceil(numericValue);
    }
  }
  return value;
};

const formatHungarianDate = (dateValue: any, showDateAndTime: boolean = false): string => {
  if (!dateValue) return '';
  
  try {
    let date: Date;
    
    // Handle Excel serial date numbers
    if (typeof dateValue === 'number' && dateValue > 1) {
      date = excelSerialToDate(dateValue);
    } else if (typeof dateValue === 'string' && dateValue.includes('.')) {
      // Parse format like "2025.06.19 7:57:20"
      const parts = dateValue.split(' ');
      const datePart = parts[0]; // "2025.06.19"
      const timePart = parts[1] || ''; // "7:57:20"
      
      const [year, month, day] = datePart.split('.').map(Number);
      const [hours = 0, minutes = 0, seconds = 0] = timePart.split(':').map(Number);
      
      // Create date (month is 0-indexed in JavaScript Date)
      date = new Date(year, month - 1, day, hours, minutes, seconds);
    } else {
      // Try standard date parsing for other formats
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) {
      return dateValue; // Return original if not a valid date
    }
    
    // Subtract 2 hours for timezone adjustment
    date.setHours(date.getHours() - 2);

    // Add 2 extra days to the display to compensate for timezones
    date.setDate(date.getDate() + 2);
    
    const options: Intl.DateTimeFormatOptions = {
      ...(showDateAndTime && {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      }),
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    };
    
    const formatted = date.toLocaleString('hu-HU', options);
    
    // For date and time, replace comma with dot to match Hungarian format
    if (showDateAndTime) {
      return formatted.replace(',', '.');
    }
    
    return formatted;
  } catch (error) {
    return dateValue; // Return original if formatting fails
  }
};

const formatTime = (time: number): number => {
  // Time i receive like 5.545454 and i want to show 5 so always round down
  return Math.floor(time);
};

const summarizeData = (data: any[][], srbnMultiplier: number, hfexMultiplier: number, individualMultipliers: { [key: string]: number } = {}): any[][] => {
  if (!data || !Array.isArray(data) || data.length === 0) return [];

  // Group data by date (index 2)
  const summaries: { [date: string]: { index3Sum: number; index4Sum: number; paymentSum: number } } = {};
  data.forEach((row, rowIndex) => {
    const date = row[2];
    if (date && typeof row[3] === 'number' && typeof row[4] === 'number') {
      if (!summaries[date]) {
        summaries[date] = { index3Sum: 0, index4Sum: 0, paymentSum: 0 };
      }
      summaries[date].index3Sum += row[3];
      summaries[date].index4Sum += row[4];
      
      // Calculate payment for this row using individual multiplier if available
      const individualMultiplier = individualMultipliers[`${rowIndex}-9`];
      const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
      summaries[date].paymentSum += payment;
    }
  });

  // Create output array with summary rows
  const result: any[][] = [];
  let currentDate: string | null = null;

  data.forEach((row, rowIndex) => {
    const date = row[2];
    
    // If the date changes or it's the last row, add a summary for the previous date
    if (date !== currentDate && currentDate && summaries[currentDate]) {
      const formattedDate = formatHungarianDate(currentDate, false);
      result.push([
        "ÖSSZEGZÉS",
        currentDate,
        null,
        `PDA Pont teljes összege: ${formatDecimal(summaries[currentDate].index3Sum.toFixed(2))}`,
        `Pont2 teljes összege: ${formatDecimal(summaries[currentDate].index4Sum.toFixed(2))}`,
        null,
        null,
        null,
        null,
        null,
        `Fizetés teljes összege: ${Math.round(summaries[currentDate].paymentSum)}`
      ]);
    }
    
    // Add the current row to the result with payment column and multiplier column
    const individualMultiplier = individualMultipliers[`${rowIndex}-9`];
    const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
    const defaultMultiplier = row[1] && row[1].toString().toUpperCase().includes('SRBN') ? srbnMultiplier : 
                             row[1] && row[1].toString().toUpperCase().includes('HF-EX') ? hfexMultiplier : 0;
    const rowWithPayment = [...row, individualMultiplier || defaultMultiplier, payment];
    result.push(rowWithPayment);
    
    currentDate = date;
  });

  // Add summary for the last date
  if (currentDate && summaries[currentDate]) {
    const formattedDate = formatHungarianDate(currentDate, false);
    result.push([
      "ÖSSZEGZÉS",
      currentDate,
      null,
      `PDA Pont teljes összege: ${summaries[currentDate].index3Sum.toFixed(2)}`,
      `Pont2 teljes összege: ${summaries[currentDate].index4Sum.toFixed(2)}`,
      null,
      null,
      null,
      null,
      null,
      `Fizetés teljes összege: ${Math.round(summaries[currentDate].paymentSum)}`
    ]);
  }

  return result;
};

const summarizeDataDetailed = (data: any[][], srbnMultiplier: number, hfexMultiplier: number, individualMultipliers: { [key: string]: number } = {}): any[][] => {
  if (!data || !Array.isArray(data) || data.length === 0) return [];

  // Group data by date (index 2), only sum index 4
  const summaries: { [date: string]: { index4Sum: number; paymentSum: number } } = {};
  data.forEach((row, rowIndex) => {
    const date = row[2];
    if (date && typeof row[4] === 'number') {
      if (!summaries[date]) {
        summaries[date] = { index4Sum: 0, paymentSum: 0 };
      }
      summaries[date].index4Sum += row[4];
      
      // Calculate payment for this row using individual multiplier if available
      const individualMultiplier = individualMultipliers[`${rowIndex}-8`];
      const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
      summaries[date].paymentSum += payment;
    }
  });

  // Create output array with summary rows
  const result: any[][] = [];
  let currentDate: string | null = null;

  data.forEach((row, rowIndex) => {
    const date = row[2];

    // If the date changes or it's the last row, add a summary for the previous date
    if (date !== currentDate && currentDate && summaries[currentDate] !== undefined) {
      result.push([
        "ÖSSZEGZÉS",
        currentDate,
        null,
        null,
        `Elszámolandó pont teljes összege: ${formatDecimal(summaries[currentDate].index4Sum.toFixed(2))}`,
        null,
        null,
        null,
        null,
        `Fizetés teljes összege: ${Math.round(summaries[currentDate].paymentSum)}`
      ]);
    }

    // Add the current row to the result with payment column and multiplier column
    const individualMultiplier = individualMultipliers[`${rowIndex}-8`];
    const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
    const defaultMultiplier = row[1] && row[1].toString().toUpperCase().includes('SRBN') ? srbnMultiplier : 
                             row[1] && row[1].toString().toUpperCase().includes('HF-EX') ? hfexMultiplier : 0;
    const rowWithPayment = [...row, individualMultiplier || defaultMultiplier, payment];
    result.push(rowWithPayment);

    currentDate = date;
  });

  // Add summary for the last date
  if (currentDate && summaries[currentDate] !== undefined) {
    result.push([
      null,
      "ÖSSZESÍTÉS",
      currentDate,
      null,
      `Elszámolandó pont teljes összege: ${formatDecimal(summaries[currentDate].index4Sum.toFixed(2))}`,
      null,
      null,
      null,
      null,
      `Fizetés teljes összege: ${Math.round(summaries[currentDate].paymentSum)}`
    ]);
  }

  return result;
};

const summarizeDataByName = (data: any[][], srbnMultiplier: number, hfexMultiplier: number, individualMultipliers: { [key: string]: number } = {}): any[][] => {
  if (!data || !Array.isArray(data) || data.length === 0) return [];

  // Group data by name (index 1)
  const summaries: { [name: string]: { index3Sum: number; index4Sum: number; paymentSum: number } } = {};
  data.forEach((row, rowIndex) => {
    const name = row[1];
    if (name && typeof row[3] === 'number' && typeof row[4] === 'number') {
      if (!summaries[name]) {
        summaries[name] = { index3Sum: 0, index4Sum: 0, paymentSum: 0 };
      }
      summaries[name].index3Sum += row[3];
      summaries[name].index4Sum += row[4];
      
      // Calculate payment for this row using individual multiplier if available
      const individualMultiplier = individualMultipliers[`${rowIndex}-9`];
      const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
      summaries[name].paymentSum += payment;
    }
  });

  // Create output array with summary rows
  const result: any[][] = [];
  let currentName: string | null = null;

  data.forEach((row, rowIndex) => {
    const name = row[1];

    // If the name changes or it's the last row, add a summary for the previous name
    if (name !== currentName && currentName && summaries[currentName]) {
      result.push([
        "ÖSSZEGZÉS",
        currentName,
        null,
        `PDA Pont teljes összege: ${formatDecimal(summaries[currentName].index3Sum.toFixed(2))}`,
        `Pont2 teljes összege: ${formatDecimal(summaries[currentName].index4Sum.toFixed(2))}`,
        null,
        null,
        null,
        null,
        null,
        `Fizetés teljes összege: ${Math.round(summaries[currentName].paymentSum)}`
      ]);
    }

    // Add the current row to the result with payment column and multiplier column
    const individualMultiplier = individualMultipliers[`${rowIndex}-9`];
    const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
    const defaultMultiplier = row[1] && row[1].toString().toUpperCase().includes('SRBN') ? srbnMultiplier : 
                             row[1] && row[1].toString().toUpperCase().includes('HF-EX') ? hfexMultiplier : 0;
    const rowWithPayment = [...row, individualMultiplier || defaultMultiplier, payment];
    result.push(rowWithPayment);

    currentName = name;
  });

  // Add summary for the last name
  if (currentName && summaries[currentName]) {
    result.push([
      "ÖSSZEGZÉS",
      currentName,
      null,
      `PDA Pont teljes összege: ${formatDecimal(summaries[currentName].index3Sum.toFixed(2))}`,
      `Pont2 teljes összege: ${formatDecimal(summaries[currentName].index4Sum.toFixed(2))}`,
      null,
      null,
      null,
      null,
      null,
      `Fizetés teljes összege: ${Math.round(summaries[currentName].paymentSum)}`
    ]);
  }

  return result;
};

const summarizeDataDetailedByName = (data: any[][], srbnMultiplier: number, hfexMultiplier: number, individualMultipliers: { [key: string]: number } = {}): any[][] => {
  if (!data || !Array.isArray(data) || data.length === 0) return [];

  // Group data by name (index 1), only sum index 4
  const summaries: { [name: string]: { index4Sum: number; paymentSum: number } } = {};
  data.forEach((row, rowIndex) => {
    const name = row[1];
    if (name && typeof row[4] === 'number') {
      if (!summaries[name]) {
        summaries[name] = { index4Sum: 0, paymentSum: 0 };
      }
      summaries[name].index4Sum += row[4];
      
      // Calculate payment for this row using individual multiplier if available
      const individualMultiplier = individualMultipliers[`${rowIndex}-8`];
      const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
      summaries[name].paymentSum += payment;
    }
  });

  // Create output array with summary rows
  const result: any[][] = [];
  let currentName: string | null = null;

  data.forEach((row, rowIndex) => {
    const name = row[1];

    // If the name changes or it's the last row, add a summary for the previous name
    if (name !== currentName && currentName && summaries[currentName] !== undefined) {
      result.push([
        "ÖSSZEGZÉS",
        currentName,
        null,
        null,
        `Elszámolandó pont teljes összege: ${formatDecimal(summaries[currentName].index4Sum.toFixed(2))}`,
        null,
        null,
        null,
        null,
        `Fizetés teljes összege: ${Math.round(summaries[currentName].paymentSum)}`
      ]);
    }

    // Add the current row to the result with payment column and multiplier column
    const individualMultiplier = individualMultipliers[`${rowIndex}-8`];
    const payment = calculatePayment(row[1], row[4], srbnMultiplier, hfexMultiplier, individualMultiplier);
    const defaultMultiplier = row[1] && row[1].toString().toUpperCase().includes('SRBN') ? srbnMultiplier : 
                             row[1] && row[1].toString().toUpperCase().includes('HF-EX') ? hfexMultiplier : 0;
    const rowWithPayment = [...row, individualMultiplier || defaultMultiplier, payment];
    result.push(rowWithPayment);

    currentName = name;
  });

  // Add summary for the last name
  if (currentName && summaries[currentName] !== undefined) {
    result.push([
      "ÖSSZEGZÉS",
      currentName,
      null,
      null,
      `Elszámolandó pont teljes összege: ${formatDecimal(summaries[currentName].index4Sum.toFixed(2))}`,
      null,
      null,
      null,
      null,
      `Fizetés teljes összege: ${Math.round(summaries[currentName].paymentSum)}`
    ]);
  }

  return result;
};

// Function to calculate payment based on name suffix and Pont 2
const calculatePayment = (name: string, pont2: number, srbnMultiplier: number, hfexMultiplier: number, individualMultiplier?: number): number => {
  if (!name || typeof pont2 !== 'number') return 0;
  
  let multiplier = 0;
  
  // Use individual multiplier if provided, otherwise use global multipliers
  if (individualMultiplier !== undefined && individualMultiplier !== null) {
    multiplier = individualMultiplier;
  } else {
    const nameStr = name.toString().toUpperCase();
    if (nameStr.includes('SRBN')) {
      multiplier = srbnMultiplier;
    } else if (nameStr.includes('HF-EX')) {
      multiplier = hfexMultiplier;
    }
  }
  
  const payment = pont2 * multiplier;
  
  // Round to 0 decimals (round 0.5 and above up)
  const decimalPart = payment % 1;
  if (decimalPart < 0.5) {
    return Math.floor(payment);
  } else {
    return Math.ceil(payment);
  }
};

// Note: excelSerialToDate function is assumed to be defined elsewhere

export default function Home() { 
  const [isDragOver, setIsDragOver] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [extractedData, setExtractedData] = useState<ExcelData[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [excelData, setExcelData] = useState<any[][]>([]);
  const [excelDataByName, setExcelDataByName] = useState<any[][]>([]);
  const [excelDataByDate, setExcelDataByDate] = useState<any[][]>([]);
  const [activeTable, setActiveTable] = useState<'A-I' | 'J-P' | 'Név szerinti összegzés' | 'Név szerinti összegzés túlóra kimutatás'>('A-I');
  const [columnsAtoI, setColumnsAtoI] = useState<any[][]>([]);
  const [columnsJtoP, setColumnsJtoP] = useState<any[][]>([]);
  const [columnsByName, setColumnsByName] = useState<any[][]>([]);
  const [columnsByNameDetailed, setColumnsByNameDetailed] = useState<any[][]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [srbnMultiplier, setSrbnMultiplier] = useState<number>(2.9);
  const [hfexMultiplier, setHfexMultiplier] = useState<number>(3.1);
  const [notes, setNotes] = useState<{ [key: string]: string }>({});
  const [multipliers, setMultipliers] = useState<{ [key: string]: number }>({});
  const [selectedRows, setSelectedRows] = useState<Set<string>>(new Set());
  const [bulkMultiplier, setBulkMultiplier] = useState<number>(0);

  const processExcelFile = useCallback((file: File) => {
    setIsProcessing(true);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheets: any[] = [];
        
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // Process the PDA pont column with rounding logic
          // const processedData = processPDAPontColumn(jsonData as any[][]);
          
          // Sort and group the data by date and name
          // const { dataByName, dataByDate } = sortAndGroupData(processedData);
          
          sheets.push({
            sheetName,
            data: jsonData
          });
        });

        setExtractedData(sheets);
      
        // Process the data from the first sheet
        if (sheets.length > 0) {
          // const processedData = processPDAPontColumn(sheets[0].data);
          // const { dataByName, dataByDate } = sortAndGroupData(processedData);

          // Extract A-I columns (columns 0-8)
          let aToIColumns = sheets[0].data.map((row: any[]) => row.slice(0, 9));
          // If the row contain 9 undefined values, remove the row
          aToIColumns = aToIColumns.filter((row: any[]) => row.some((cell: any) => cell !== undefined));
          // Remove empty arrays from aToIColumns
          aToIColumns = aToIColumns.filter((row: any[]) => row.length > 0);
          aToIColumns = aToIColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
          // Additional filter to remove rows with only null/empty values
          aToIColumns = aToIColumns.filter((row: any[]) => {
            const meaningfulCells = row.filter((cell: any) => 
              cell !== null && cell !== undefined && cell !== '' && cell !== 0
            );
            return meaningfulCells.length > 0;
          });

          aToIColumns = aToIColumns.sort((a: any[], b: any[]) => {
            const dateAStr = a[2];
            const dateBStr = b[2];
            
            // Handle Hungarian date format "2025.06.24"
            if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
              // Parse "2025.06.24" format
              const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
              const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
              
              const dateA = new Date(yearA, monthA - 1, dayA); // month is 0-indexed
              const dateB = new Date(yearB, monthB - 1, dayB);
              
              return dateA.getTime() - dateB.getTime(); // Sort ascending (oldest first)
            }
            
            return 0; // If dates are invalid, keep original order
          });
          aToIColumns = aToIColumns.slice(1);

          // aToIColumns = aToIColumns.filter((row: any[]) => row[0] !== `=== undefined - DAILY SUMMARY ===`);
          aToIColumns = summarizeData(aToIColumns, srbnMultiplier, hfexMultiplier, multipliers);
          setColumnsAtoI(aToIColumns);
          
          // Extract J-P columns (columns 9-15)
          let jToPColumns = sheets[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);
          // Filter out empty rows and rows with undefined values
          jToPColumns = jToPColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
            jToPColumns = jToPColumns.sort((a: any[], b: any[]) => {
             const dateAStr = a[2];
             const dateBStr = b[2];
             
             // Handle Hungarian date format "2025.06.24"
             if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
               // Parse "2025.06.24" format
               const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
               const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
               
               const dateA = new Date(yearA, monthA - 1, dayA); // month is 0-indexed
               const dateB = new Date(yearB, monthB - 1, dayB);
               
               return dateA.getTime() - dateB.getTime(); // Sort ascending (oldest first)
             }
             
             return 0; // If dates are invalid, keep original order
           });
          
          jToPColumns = summarizeDataDetailed(jToPColumns, srbnMultiplier, hfexMultiplier, multipliers);
          setColumnsJtoP(jToPColumns);

          let columnsByName = sheets[0].data.map((row: any[]) => row.slice(0, 9));

          // Sort by name (index 1), safely handle undefined/null

          // Remove the empty arrays from columnsByName
          columnsByName = columnsByName
            .filter((row: any[]) => row.length > 0)
            .filter((row: any[]) => row.some((cell: any) => cell !== null && cell !== undefined && cell !== ''))
            .filter((row: any[]) => {
              // Check if the row has meaningful data (not just nulls and empty values)
              const meaningfulCells = row.filter((cell: any) => 
                cell !== null && cell !== undefined && cell !== '' && cell !== 0
              );
              return meaningfulCells.length > 0;
            });


          columnsByName.sort((a: any[], b: any[]) => {
            const nameA = (a[1] ?? '').toString();
            const nameB = (b[1] ?? '').toString();
            return nameA.localeCompare(nameB);
          });
          columnsByName = summarizeDataByName(columnsByName, srbnMultiplier, hfexMultiplier, multipliers);

          setColumnsByName(columnsByName);

          let columnsByNameDetailed = sheets[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);


          columnsByNameDetailed = summarizeDataDetailedByName(columnsByNameDetailed, srbnMultiplier, hfexMultiplier, multipliers);
          setColumnsByNameDetailed(columnsByNameDetailed);

          
          setExcelData(sheets[0].data);
          setExcelDataByName(sheets[0].data);
          setExcelDataByDate(sheets[0].data);
          setFileName(file.name);
          setError(null);
        }
      } catch (error) {
        console.error('Error processing Excel file:', error);
        alert('Hiba történt az Excel fájl feldolgozása során. Kérjük, ellenőrizd, hogy érvényes Excel fájlt választottál-e.');
      } finally {
        setIsProcessing(false);
      }
    };

    reader.onerror = () => {
      console.error('Error reading file');
      alert('Hiba történt a fájl olvasása során');
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const handleFileSelect = useCallback((files: FileList | null) => {
    if (!files || files.length === 0) return;
    
    const file = files[0];
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
    
    if (!validExtensions.includes(fileExtension)) {
      alert('Kérjük, válassz érvényes Excel fájlt (.xlsx vagy .xls)');
      return;
    }

    processExcelFile(file);
  }, [processExcelFile]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    handleFileSelect(e.dataTransfer.files);
  }, [handleFileSelect]);

  const handleFileInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    handleFileSelect(e.target.files);
  }, [handleFileSelect]);

  const handleBrowseClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleNoteChange = useCallback((rowKey: string, value: string) => {
    setNotes(prev => ({
      ...prev,
      [rowKey]: value
    }));
  }, []);

  const handleMultiplierChange = useCallback((rowKey: string, value: number) => {
    setMultipliers(prev => ({
      ...prev,
      [rowKey]: value
    }));
  }, []);

  const handleRowSelection = useCallback((rowKey: string, isSelected: boolean) => {
    setSelectedRows(prev => {
      const newSet = new Set(prev);
      if (isSelected) {
        newSet.add(rowKey);
      } else {
        newSet.delete(rowKey);
      }
      return newSet;
    });
  }, []);

  const handleSelectAll = useCallback((isSelected: boolean, tableType: string) => {
    if (isSelected) {
      // Get all row keys for the current table
      let data: any[][] = [];
      let multiplierColumnIndex = 9; // Default for A-I tables
      
      switch (tableType) {
        case 'A-I':
          data = columnsAtoI.slice(1);
          multiplierColumnIndex = 9;
          break;
        case 'J-P':
          data = columnsJtoP;
          multiplierColumnIndex = 8;
          break;
        case 'Név szerinti összegzés':
          data = columnsByName;
          multiplierColumnIndex = 9;
          break;
        case 'Név szerinti összegzés túlóra kimutatás':
          data = columnsByNameDetailed;
          multiplierColumnIndex = 8;
          break;
      }
      
      const allRowKeys = data
        .filter((row: any[]) => row && row.some(cell => cell !== null && cell !== undefined && cell !== ''))
        .filter((row: any[]) => !(row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS')))
        .map((_, index) => `${index}-${multiplierColumnIndex}`);
      
      setSelectedRows(new Set(allRowKeys));
    } else {
      setSelectedRows(new Set());
    }
  }, [columnsAtoI, columnsJtoP, columnsByName, columnsByNameDetailed]);

  const handleBulkMultiplierChange = useCallback(() => {
    if (selectedRows.size === 0 || !bulkMultiplier) return;
    
    const updates: { [key: string]: number } = {};
    selectedRows.forEach(rowKey => {
      updates[rowKey] = bulkMultiplier;
    });
    
    setMultipliers(prev => ({
      ...prev,
      ...updates
    }));
    
    // Clear selection after bulk update
    setSelectedRows(new Set());
    setBulkMultiplier(0);
  }, [selectedRows, bulkMultiplier]);

  const recalculatePayments = useCallback(() => {
    if (extractedData.length > 0 && extractedData[0].data.length > 0) {
      // Recalculate A-I columns
      let aToIColumns = extractedData[0].data.map((row: any[]) => row.slice(0, 9));
      aToIColumns = aToIColumns.filter((row: any[]) => row.some((cell: any) => cell !== undefined));
      aToIColumns = aToIColumns.filter((row: any[]) => row.length > 0);
      aToIColumns = aToIColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
      // Additional filter to remove rows with only null/empty values
      aToIColumns = aToIColumns.filter((row: any[]) => {
        const meaningfulCells = row.filter((cell: any) => 
          cell !== null && cell !== undefined && cell !== '' && cell !== 0
        );
        return meaningfulCells.length > 0;
      });

      aToIColumns = aToIColumns.sort((a: any[], b: any[]) => {
        const dateAStr = a[2];
        const dateBStr = b[2];
        
        if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
          const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
          const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
          
          const dateA = new Date(yearA, monthA - 1, dayA);
          const dateB = new Date(yearB, monthB - 1, dayB);
          
          return dateA.getTime() - dateB.getTime();
        }
        
        return 0;
      });

      aToIColumns = summarizeData(aToIColumns, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsAtoI(aToIColumns);

      // Recalculate J-P columns
      let jToPColumns = extractedData[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);
      jToPColumns = jToPColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
      jToPColumns = jToPColumns.sort((a: any[], b: any[]) => {
        const dateAStr = a[2];
        const dateBStr = b[2];
        
        if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
          const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
          const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
          
          const dateA = new Date(yearA, monthA - 1, dayA);
          const dateB = new Date(yearB, monthB - 1, dayB);
          
          return dateA.getTime() - dateB.getTime();
        }
        
        return 0;
      });
      jToPColumns = summarizeDataDetailed(jToPColumns, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsJtoP(jToPColumns);

      // Recalculate name-based columns
      let columnsByName = extractedData[0].data.map((row: any[]) => row.slice(0, 9));
      
      // Remove the empty arrays from columnsByName
      columnsByName = columnsByName
        .filter((row: any[]) => row.length > 0)
        .filter((row: any[]) => row.some((cell: any) => cell !== null && cell !== undefined && cell !== ''))
        .filter((row: any[]) => {
          // Check if the row has meaningful data (not just nulls and empty values)
          const meaningfulCells = row.filter((cell: any) => 
            cell !== null && cell !== undefined && cell !== '' && cell !== 0
          );
          return meaningfulCells.length > 0;
        });
      
      columnsByName.sort((a: any[], b: any[]) => {
        const nameA = (a[1] ?? '').toString();
        const nameB = (b[1] ?? '').toString();
        return nameA.localeCompare(nameB);
      });
      columnsByName = summarizeDataByName(columnsByName, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsByName(columnsByName);

      // Recalculate name-based detailed columns
      let columnsByNameDetailed = extractedData[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);
      columnsByNameDetailed = summarizeDataDetailedByName(columnsByNameDetailed, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsByNameDetailed(columnsByNameDetailed);
    }
  }, [extractedData, srbnMultiplier, hfexMultiplier, multipliers]);

  // Recalculate data when multipliers change
  useEffect(() => {
    if (extractedData.length > 0 && extractedData[0].data.length > 0) {
      // Recalculate A-I columns
      let aToIColumns = extractedData[0].data.map((row: any[]) => row.slice(0, 9));
      aToIColumns = aToIColumns.filter((row: any[]) => row.some((cell: any) => cell !== undefined));
      aToIColumns = aToIColumns.filter((row: any[]) => row.length > 0);
      aToIColumns = aToIColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
      // Additional filter to remove rows with only null/empty values
      aToIColumns = aToIColumns.filter((row: any[]) => {
        const meaningfulCells = row.filter((cell: any) => 
          cell !== null && cell !== undefined && cell !== '' && cell !== 0
        );
        return meaningfulCells.length > 0;
      });

      aToIColumns = aToIColumns.sort((a: any[], b: any[]) => {
        const dateAStr = a[2];
        const dateBStr = b[2];
        
        if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
          const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
          const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
          
          const dateA = new Date(yearA, monthA - 1, dayA);
          const dateB = new Date(yearB, monthB - 1, dayB);
          
          return dateA.getTime() - dateB.getTime();
        }
        
        return 0;
      });

      aToIColumns = summarizeData(aToIColumns, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsAtoI(aToIColumns);

      // Recalculate J-P columns
      let jToPColumns = extractedData[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);
      jToPColumns = jToPColumns.filter((row: any[]) => row && row.length > 0 && row.some((cell: any) => cell !== null && cell !== undefined));
      jToPColumns = jToPColumns.sort((a: any[], b: any[]) => {
        const dateAStr = a[2];
        const dateBStr = b[2];
        
        if (dateAStr && dateBStr && typeof dateAStr === 'string' && typeof dateBStr === 'string') {
          const [yearA, monthA, dayA] = dateAStr.split('.').map(Number);
          const [yearB, monthB, dayB] = dateBStr.split('.').map(Number);
          
          const dateA = new Date(yearA, monthA - 1, dayA);
          const dateB = new Date(yearB, monthB - 1, dayB);
          
          return dateA.getTime() - dateB.getTime();
        }
        
        return 0;
      });
      jToPColumns = summarizeDataDetailed(jToPColumns, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsJtoP(jToPColumns);

      // Recalculate name-based columns
      let columnsByName = extractedData[0].data.map((row: any[]) => row.slice(0, 9));
      
      // Remove the empty arrays from columnsByName
      columnsByName = columnsByName
        .filter((row: any[]) => row.length > 0)
        .filter((row: any[]) => row.some((cell: any) => cell !== null && cell !== undefined && cell !== ''))
        .filter((row: any[]) => {
          // Check if the row has meaningful data (not just nulls and empty values)
          const meaningfulCells = row.filter((cell: any) => 
            cell !== null && cell !== undefined && cell !== '' && cell !== 0
          );
          return meaningfulCells.length > 0;
        });
      
      columnsByName.sort((a: any[], b: any[]) => {
        const nameA = (a[1] ?? '').toString();
        const nameB = (b[1] ?? '').toString();
        return nameA.localeCompare(nameB);
      });
      columnsByName = summarizeDataByName(columnsByName, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsByName(columnsByName);

      // Recalculate name-based detailed columns
      let columnsByNameDetailed = extractedData[0].data.map((row: any[]) => row.slice(9, 16)).slice(1);
      columnsByNameDetailed = summarizeDataDetailedByName(columnsByNameDetailed, srbnMultiplier, hfexMultiplier, multipliers);
      setColumnsByNameDetailed(columnsByNameDetailed);
    }
  }, [srbnMultiplier, hfexMultiplier, multipliers, extractedData]);

  const clearData = useCallback(() => {
    setExtractedData([]);
    setExcelData([]);
    setExcelDataByName([]);
    setExcelDataByDate([]);
    setColumnsAtoI([]);
    setColumnsJtoP([]);
    setActiveTable('A-I');
    setFileName('');
    setSrbnMultiplier(2);
    setHfexMultiplier(3);
    setNotes({});
    setMultipliers({});
    setSelectedRows(new Set());
    setBulkMultiplier(0);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }, []);

  const downloadTable = async (type: 'A-I' | 'J-P' | 'Név szerinti összegzés' | 'Név szerinti összegzés túlóra kimutatás') => {
    const workbook = new ExcelJS.Workbook();
    
    const headers = [
      'Szedőkód', 'Szedő neve', 'Dátum', 'PDA Pont', 'Pont 2',
      'Műszak kezdete', 'Műszak vége', 'Összesített óra', 'Megjegyzés', 'Szorzó', 'Fizetés'
    ];
    const headersDetailed = [
      'Szedőkód', 'Szedő neve', 'Dátum', 'Túlóra', 'Elszámolandó pont',
      'Műszak kezdet', 'Műszak vége', 'Megjegyzés', 'Szorzó', 'Fizetés'
    ];

    let table: any[][];
    if (type === 'A-I') {
      table = columnsAtoI.slice(1);
    } else if (type === 'J-P') {
      table = columnsJtoP.slice(1);
    } else if (type === 'Név szerinti összegzés') {
      table = columnsByName.slice(1).filter(
        (row: any[]) => Array.isArray(row) && row.length > 0 && row.some((cell: any) => cell !== undefined && cell !== null && cell !== '')
      );
    } else {
      table = columnsByNameDetailed.slice(1);
    }

    // Helper to add notes and multipliers to rows
    const addNotesToRows = (rows: any[][], startIndex: number = 0, isDetailed: boolean = false) => {
      return rows.map((row, idx) => {
        const rowWithNotes = [...row];
        if (isDetailed) {
          // For detailed tables: notes at index 7, multipliers at index 8
          if (rowWithNotes.length >= 7) {
            const noteKey = `${startIndex + idx}-7`;
            rowWithNotes[7] = notes[noteKey] || rowWithNotes[7] || '';
          }
          if (rowWithNotes.length >= 8) {
            const multiplierKey = `${startIndex + idx}-8`;
            rowWithNotes[8] = multipliers[multiplierKey] || rowWithNotes[8] || '';
          }
        } else {
          // For standard tables: notes at index 8, multipliers at index 9
          if (rowWithNotes.length >= 8) {
            const noteKey = `${startIndex + idx}-8`;
            rowWithNotes[8] = notes[noteKey] || rowWithNotes[8] || '';
          }
          if (rowWithNotes.length >= 9) {
            const multiplierKey = `${startIndex + idx}-9`;
            rowWithNotes[9] = multipliers[multiplierKey] || rowWithNotes[9] || '';
          }
        }
        return rowWithNotes;
      });
    };

    // Helper to add a worksheet for a group
    const addSheet = (sheetName: string, header: string[], rows: any[][]) => {
      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.addRow(header);
      worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getRow(1).height = 24;
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '3B82F6' },
      };
      rows.forEach((row, idx) => {
        const addedRow = worksheet.addRow(row);
        if (row.some((cell: any) => typeof cell === 'string' && cell.includes('ÖSSZEGZÉS'))) {
          addedRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '38b538' },
            };
          });
        } else {
          if (idx % 2 === 0) {
            addedRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E0ECFB' },
            };
          }
        }
        // Set number format to '0' (no decimals) for all cells
        addedRow.eachCell({ includeEmpty: true }, (cell) => {
          if (typeof cell.value === 'number') {
            cell.numFmt = '0';
            cell.alignment = { horizontal: 'right' };
          }
        });

        // Format columns 6-7 (Műszak kezdete and Műszak vége) as time
        const startTimeCell = addedRow.getCell(6);
        const endTimeCell = addedRow.getCell(7);
        if (startTimeCell.value && typeof startTimeCell.value === 'number') {
          startTimeCell.value = formatTime(startTimeCell.value);
          startTimeCell.alignment = { horizontal: 'center' };
        }
        if (endTimeCell.value && typeof endTimeCell.value === 'number') {
          endTimeCell.value = formatTime(endTimeCell.value);
          endTimeCell.alignment = { horizontal: 'center' };
        }
        
        // Format Dátum column (index 3) - keep date format
        const dateCell = addedRow.getCell(3);
        if (dateCell.value) {
          dateCell.numFmt = 'yyyy.mm.dd';
          dateCell.alignment = { horizontal: 'center' };
        }
      });
      worksheet.columns.forEach((column: any) => {
        let maxLength = 10;
        column.eachCell({ includeEmpty: true }, (cell: any) => {
          const length = cell.value ? cell.value.toString().length : 0;
          maxLength = Math.max(maxLength, length);
        });
        column.width = maxLength + 2;
      });
      worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    };

    if (type === 'A-I' || type === 'J-P') {
      // Group by date (index 2)
      const groupByDate: { [date: string]: any[][] } = {};
      table.forEach(row => {
        const date = row[2];
        if (!date) return;
        if (!groupByDate[date]) groupByDate[date] = [];
        groupByDate[date].push(row);
      });
      // For each date, add a worksheet
      Object.entries(groupByDate).forEach(([date, rows]) => {
        // Find summary row(s) for this date
        const summaryRows = rows.filter(row => row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS'));
        const dataRows = rows.filter(row => !(row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS')));
        
        // Add notes to data rows
        const isDetailed = type === 'J-P';
        const dataRowsWithNotes = addNotesToRows(dataRows, 0, isDetailed);
        const summaryRowsWithNotes = addNotesToRows(summaryRows, dataRows.length, isDetailed);
        
        addSheet(date, isDetailed ? headersDetailed : headers, [...dataRowsWithNotes, ...summaryRowsWithNotes]);
      });
    } else if (type === 'Név szerinti összegzés' || type === 'Név szerinti összegzés túlóra kimutatás') {
      // Group by name (index 1)
      const groupByName: { [name: string]: any[][] } = {};
      table.forEach(row => {
        const name = row[1];
        if (!name) return;
        if (!groupByName[name]) groupByName[name] = [];
        groupByName[name].push(row);
      });
      Object.entries(groupByName).forEach(([name, rows]) => {
        // Find summary row(s) for this name
        const summaryRows = rows.filter(row => row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS'));
        const dataRows = rows.filter(row => !(row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS')));
        
        // Add notes to data rows
        const isDetailed = type === 'Név szerinti összegzés túlóra kimutatás';
        const dataRowsWithNotes = addNotesToRows(dataRows, 0, isDetailed);
        const summaryRowsWithNotes = addNotesToRows(summaryRows, dataRows.length, isDetailed);
        
        addSheet(name, isDetailed ? headersDetailed : headers, [...dataRowsWithNotes, ...summaryRowsWithNotes]);
      });
    }

    // Remove the default worksheet (created at start)
    if (workbook.worksheets.length > 1) {
      workbook.removeWorksheet('Export');
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `${type}.xlsx`);
  };
  

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 dark:from-gray-900 dark:to-gray-800">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          {/* Header */}
          <div className="text-center mb-8">
            <h1 className="text-4xl font-bold text-gray-900 dark:text-white mb-4">
              Excel Fájl Kibontó
            </h1>
            <p className="text-lg text-gray-600 dark:text-gray-300">
              Töltsd fel Excel fájljait és nyerd ki a tartalmukat
            </p>
          </div>

          {/* Upload Area */}
          <div className="bg-white dark:bg-gray-800 rounded-lg shadow-lg p-8 mb-8">
            <div
              className={`border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200 ${
                isDragOver
                  ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20'
                  : 'border-gray-300 dark:border-gray-600 hover:border-gray-400 dark:hover:border-gray-500'
              }`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="space-y-4">
                <div className="text-6xl text-gray-400 dark:text-gray-500">
                  📊
                </div>
                <div>
                  <h3 className="text-xl font-semibold text-gray-900 dark:text-white mb-2">
                    {isDragOver ? 'Engedd el az Excel fájlt ide' : 'Excel Fájl Feltöltése'}
                  </h3>
                  <p className="text-gray-600 dark:text-gray-300 mb-4">
                    Húzd és engedd el a .xlsx vagy .xls fájlt ide, vagy kattints a böngészéshez
                  </p>
                </div>
                
                {/* Input Fields */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-w-2xl mx-auto">
                  
                  <div>
                    <label htmlFor="srbnMultiplier" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                      SRBN Szorzó
                    </label>
                    <input
                      type="number"
                      id="srbnMultiplier"
                      value={srbnMultiplier}
                      onChange={(e) => setSrbnMultiplier(Number(e.target.value))}
                      placeholder="SRBN"
                      className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                    />
                  </div>
                  
                  <div>
                    <label htmlFor="hfexMultiplier" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                      HFEX Szorzó
                    </label>
                    <input
                      type="number"
                      id="hfexMultiplier"
                      value={hfexMultiplier}
                      onChange={(e) => setHfexMultiplier(Number(e.target.value))}
                      placeholder="HF-EX"
                      className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                    />
                  </div>
                </div>
                
                <div className="flex flex-col sm:flex-row gap-4 justify-center">
                  <button
                    onClick={handleBrowseClick}
                    disabled={isProcessing}
                    className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-medium rounded-lg transition-colors duration-200"
                  >
                    {isProcessing ? 'Feldolgozás...' : 'Fájlok Böngészése'}
                  </button>
                  
                  {extractedData.length > 0 && (
                    <button
                      onClick={clearData}
                      className="px-6 py-3 bg-gray-600 hover:bg-gray-700 text-white font-medium rounded-lg transition-colors duration-200"
                    >
                      Adatok Törlése
                    </button>
                  )}
                </div>
                
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileInputChange}
                  className="hidden"
                />
              </div>
            </div>
          </div>

          {/* Processing Indicator */}
          {isProcessing && (
            <div className="bg-white dark:bg-gray-800 rounded-lg shadow-lg p-6 mb-8">
              <div className="flex items-center justify-center space-x-3">
                <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
                <span className="text-gray-700 dark:text-gray-300">Excel fájl feldolgozása...</span>
              </div>
            </div>
          )}

          {/* Results */}
          {(columnsAtoI.length > 0 || columnsJtoP.length > 0) && (
            <div className="space-y-6">
              <div className="bg-white dark:bg-gray-800 rounded-lg shadow-lg p-6">
                <h2 className="text-2xl font-bold text-gray-900 dark:text-white mb-4">
                  Kibontott Adatok
                </h2>
                <div className="mb-4">
                  <p className="text-gray-600 dark:text-gray-300">
                    <span className="font-semibold">Fájl:</span> {fileName}
                  </p>
                  <p className="text-gray-600 dark:text-gray-300">
                    <span className="font-semibold">Lapok:</span> {extractedData.length}
                  </p>
                </div>
                
                <div className="bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-800 rounded-lg p-4">
                  <div className="flex items-center space-x-2">
                    <div className="text-green-600 dark:text-green-400">✓</div>
                    <span className="text-green-800 dark:text-green-200 font-medium">
                      Adatok sikeresen kibontva! Nézd át a táblázatokat és tölsd le a számodra tökéletesen megfelelő fájlt!
                    </span>
                  </div>
                </div>
              </div>

              {/* Sheet Tabs */}
              <div className="bg-white dark:bg-gray-800 rounded-lg shadow-lg">
                <div className="border-b border-gray-200 dark:border-gray-700">
                  <nav className="flex space-x-8 px-6" aria-label="Tabs">
                    {extractedData.map((sheet, index) => (
                      <button
                        key={index}
                        className="py-4 px-1 border-b-2 border-transparent text-sm font-medium text-gray-500 hover:text-gray-700 hover:border-gray-300 dark:text-gray-400 dark:hover:text-gray-300 dark:hover:border-gray-600"
                      >
                        {sheet.sheetName} ({sheet.data.length} sor)
                      </button>
                    ))}
                  </nav>
                </div>

                {/* Table Tabs */}
                <div className="border-b border-gray-200 dark:border-gray-700">
                  <nav className="flex space-x-8 px-6" aria-label="Tabs">
                    <button
                      className={`${activeTable === 'A-I' ? 'border-blue-500 dark:border-blue-400' : ''} py-4 px-1 border-b-2 border-transparent text-sm font-medium text-gray-500 hover:text-gray-700 hover:border-gray-300 dark:text-gray-400 dark:hover:text-gray-300 dark:hover:border-gray-600`}
                      onClick={() => setActiveTable('A-I')}
                      disabled={isLoading}
                    > 
                      Összegzés
                    </button>
                    <button
                      className={`${activeTable === 'J-P' ? 'border-blue-500 dark:border-blue-400' : ''} py-4 px-1 border-b-2 border-transparent text-sm font-medium text-gray-500 hover:text-gray-700 hover:border-gray-300 dark:text-gray-400 dark:hover:text-gray-300 dark:hover:border-gray-600`}
                      onClick={() => setActiveTable('J-P')}
                      disabled={isLoading}
                    >
                      Túlóra kimutatás
                    </button>
                    <button
                      className={`${activeTable === 'Név szerinti összegzés' ? 'border-blue-500 dark:border-blue-400' : ''} py-4 px-1 border-b-2 border-transparent text-sm font-medium text-gray-500 hover:text-gray-700 hover:border-gray-300 dark:text-gray-400 dark:hover:text-gray-300 dark:hover:border-gray-600`}
                      onClick={() => setActiveTable('Név szerinti összegzés')}
                      disabled={isLoading}
                    >
                      Név szerinti összegzés
                    </button>
                    <button
                      className={`${activeTable === 'Név szerinti összegzés túlóra kimutatás' ? 'border-blue-500 dark:border-blue-400' : ''} py-4 px-1 border-b-2 border-transparent text-sm font-medium text-gray-500 hover:text-gray-700 hover:border-gray-300 dark:text-gray-400 dark:hover:text-gray-300 dark:hover:border-gray-600`}
                      onClick={() => setActiveTable('Név szerinti összegzés túlóra kimutatás')}
                      disabled={isLoading}
                    >
                      Név szerinti összegzés túlóra kimutatás
                    </button>
                  </nav>
                </div>  

                {/* Sheet Content */}
                <div className="p-6">
                  {extractedData.map((sheet, sheetIndex) => (
                    <div key={sheetIndex} className="space-y-8">
                      <h3 className="text-lg font-semibold text-gray-900 dark:text-white">
                        <div className="flex gap-3">
                          <button 
                          onClick={() => downloadTable(activeTable)}
                          className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-medium rounded-lg transition-colors duration-200"
                          >
                            {activeTable === 'A-I' ? 'Szimpla összegzés' : activeTable === 'J-P' ? 'Túlóra kimutatás alapú összegzés' : activeTable === 'Név szerinti összegzés' ? 'Név szerinti összegzés' : 'Név szerinti összegzés túlóra kimutatás'} letöltése
                          </button>
                          <button 
                          onClick={recalculatePayments}
                          className="px-6 py-3 bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white font-medium rounded-lg transition-colors duration-200"
                          title="Fizetések újraszámítása a jelenlegi szorzók alapján"
                          >
                            🔄 Fizetés újraszámítása
                          </button>
                        </div>
                      </h3>

                      {/* Bulk Edit Interface */}
                      {selectedRows.size > 0 && (
                        <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg p-4 mb-4">
                          <div className="flex items-center gap-4">
                            <span className="text-blue-800 dark:text-blue-200 font-medium">
                              {selectedRows.size} sor kiválasztva
                            </span>
                            <div className="flex items-center gap-2">
                              <label htmlFor="bulkMultiplier" className="text-sm font-medium text-blue-700 dark:text-blue-300">
                                Új szorzó:
                              </label>
                              <input
                                id="bulkMultiplier"
                                type="number"
                                step="0.1"
                                value={bulkMultiplier || ''}
                                onChange={(e) => setBulkMultiplier(Number(e.target.value))}
                                placeholder="Szorzó érték"
                                className="w-24 px-2 py-1 text-sm border border-blue-300 dark:border-blue-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                              />
                              <button
                                onClick={handleBulkMultiplierChange}
                                disabled={!bulkMultiplier}
                                className="px-4 py-1 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white text-sm font-medium rounded transition-colors duration-200"
                              >
                                Alkalmaz
                              </button>
                              <button
                                onClick={() => setSelectedRows(new Set())}
                                className="px-4 py-1 bg-gray-600 hover:bg-gray-700 text-white text-sm font-medium rounded transition-colors duration-200"
                              >
                                Kijelölés törlése
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                      
                      {/* A-I Table */}
                      {activeTable === 'A-I' && columnsAtoI.length > 0 && (
                        <div className="space-y-4">
                          <h4 className="text-lg font-semibold text-gray-900 dark:text-white">
                            Szimpla összegzés
                          </h4>
                          <div className="overflow-x-auto">
                            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700">
                              <thead className="bg-gray-50 dark:bg-gray-700">
                                <tr>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    <input
                                      type="checkbox"
                                      onChange={(e) => handleSelectAll(e.target.checked, 'A-I')}
                                      className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                    />
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedőkód
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedő neve
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Dátum
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    PDA Pont
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Pont 2
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak kezdete
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak vége
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Összesített óra
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Megjegyzés
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szorzó
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Fizetés
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-700">
                                {columnsAtoI
                                  .slice(1)
                                  .filter((row: any[]) => {
                                    // Filter out completely empty rows
                                    return row && row.some(cell => 
                                      cell !== null && cell !== undefined && cell !== ''
                                    );
                                  })
                                  .map((row: any[], rowIndex: number) => {
                                  // Check if this is a summary row
                                  const isSummaryRow = row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS');
                                  
                                  return (
                                    <tr 
                                      key={rowIndex} 
                                      className={`${
                                        isSummaryRow 
                                          ? 'bg-blue-50 dark:bg-blue-900/20 font-semibold border-t-2 border-blue-200 dark:border-blue-700' 
                                          : 'hover:bg-gray-50 dark:hover:bg-gray-700'
                                      }`}
                                    >
                                      <td className="px-6 py-4 whitespace-nowrap text-sm">
                                        {!isSummaryRow && (
                                          <input
                                            type="checkbox"
                                            checked={selectedRows.has(`${rowIndex}-9`)}
                                            onChange={(e) => handleRowSelection(`${rowIndex}-9`, e.target.checked)}
                                            className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                          />
                                        )}
                                      </td>
                                      {row.map((cell: any, cellIndex: number) => (
                                        <td
                                          key={cellIndex}
                                          className={`px-6 py-4 whitespace-nowrap text-sm ${
                                            isSummaryRow 
                                              ? 'text-blue-800 dark:text-blue-200 font-medium' 
                                              : 'text-gray-900 dark:text-gray-300'
                                          }`}
                                        >
                                          {cellIndex === 8 ? (
                                            <input
                                              type="text"
                                              value={notes[`${rowIndex}-${cellIndex}`] || ''}
                                              onChange={(e) => handleNoteChange(`${rowIndex}-${cellIndex}`, e.target.value)}
                                              placeholder="Megjegyzés..."
                                              className="w-full min-w-[120px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 9 ? (
                                            <input
                                              type="number"
                                              step="0.1"
                                              value={multipliers[`${rowIndex}-${cellIndex}`] || cell || ''}
                                              onChange={(e) => handleMultiplierChange(`${rowIndex}-${cellIndex}`, Number(e.target.value))}
                                              placeholder="Szorzó"
                                              className="w-full min-w-[80px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 6 || cellIndex === 5 ? formatHungarianDate(cell) : cellIndex === 3 || cellIndex === 4 || cellIndex === 10 ? formatDecimal(cell) : cellIndex === 7 ? formatTime(cell) : cell || ''}
                                          {cellIndex === 10 ? ' Ft' : ''}
                                        </td>
                                      ))}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {/* J-P Table */}
                      {activeTable === 'J-P' && columnsJtoP.length > 0 && (
                        <div className="space-y-4">
                          <h4 className="text-lg font-semibold text-gray-900 dark:text-white">
                            Túlóra kimutatás alapú összegzés
                          </h4>
                          <div className="overflow-x-auto">
                            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700">
                              <thead className="bg-gray-50 dark:bg-gray-700">
                                <tr>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    <input
                                      type="checkbox"
                                      onChange={(e) => handleSelectAll(e.target.checked, 'J-P')}
                                      className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                    />
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedőkód
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedő neve
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Dátum
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Túlóra
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    PDA Pont
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak kezdete
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak vége
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Megjegyzés
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szorzó
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Fizetés
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-700">
                                {columnsJtoP
                                  .filter((row: any[]) => {
                                    // Filter out completely empty rows
                                    return row && row.some(cell => 
                                      cell !== null && cell !== undefined && cell !== ''
                                    );
                                  })
                                  .map((row: any[], rowIndex: number) => {
                                  // Check if this is a summary row
                                  const isSummaryRow = row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS');
                                  
                                  return (
                                    <tr 
                                      key={rowIndex} 
                                      className={`${
                                        isSummaryRow 
                                          ? 'bg-blue-50 dark:bg-blue-900/20 font-semibold border-t-2 border-blue-200 dark:border-blue-700' 
                                          : 'hover:bg-gray-50 dark:hover:bg-gray-700'
                                      }`}
                                    >
                                      <td className="px-6 py-4 whitespace-nowrap text-sm">
                                        {!isSummaryRow && (
                                          <input
                                            type="checkbox"
                                            checked={selectedRows.has(`${rowIndex}-8`)}
                                            onChange={(e) => handleRowSelection(`${rowIndex}-8`, e.target.checked)}
                                            className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                          />
                                        )}
                                      </td>
                                      {row.map((cell: any, cellIndex: number) => (
                                        <td
                                          key={cellIndex}
                                          className={`px-6 py-4 whitespace-nowrap text-sm ${
                                            isSummaryRow 
                                              ? 'text-blue-800 dark:text-blue-200 font-medium' 
                                              : 'text-gray-900 dark:text-gray-300'
                                          }`}
                                        >
                                          {cellIndex === 7 ? (
                                            <input
                                              type="text"
                                              value={notes[`${rowIndex}-${cellIndex}`] || ''}
                                              onChange={(e) => handleNoteChange(`${rowIndex}-${cellIndex}`, e.target.value)}
                                              placeholder="Megjegyzés..."
                                              className="w-full min-w-[120px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 8 ? (
                                            <input
                                              type="number"
                                              step="0.1"
                                              value={multipliers[`${rowIndex}-${cellIndex}`] || cell || ''}
                                              onChange={(e) => handleMultiplierChange(`${rowIndex}-${cellIndex}`, Number(e.target.value))}
                                              placeholder="Szorzó"
                                              className="w-full min-w-[80px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 9 ? formatDecimal(cell) : cellIndex === 3 || cellIndex === 4 ? formatDecimal(cell) : cellIndex === 5 || cellIndex === 6 ? formatHungarianDate(cell, true) : cell || ''}
                                          {cellIndex === 9 ? ' Ft' : ''}
                                          </td>
                                      ))}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {activeTable === 'Név szerinti összegzés' && columnsByName.length > 0 && (
                        <div className="space-y-4">
                          <h4 className="text-lg font-semibold text-gray-900 dark:text-white">
                            Név szerinti összegzés
                          </h4>
                          <div className="overflow-x-auto">
                            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700">
                              <thead className="bg-gray-50 dark:bg-gray-700">
                                <tr>
                                  {/* {columnsByName[0]?.map((cell: any, cellIndex: number) => (
                                    <th
                                      key={cellIndex}
                                      className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider"
                                    >
                                      {cell}
                                    </th>
                                  ))} */}
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    <input
                                      type="checkbox"
                                      onChange={(e) => handleSelectAll(e.target.checked, 'Név szerinti összegzés')}
                                      className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                    />
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedőkód
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedő neve
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Dátum
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    PDA Pont
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Pont 2
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak kezdete
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak vége
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Összesített óra
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Megjegyzés
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szorzó
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Fizetés
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-700">
                                {columnsByName
                                  .filter((row: any[]) => {
                                    // Filter out completely empty rows
                                    return row && row.some(cell => 
                                      cell !== null && cell !== undefined && cell !== ''
                                    );
                                  })
                                  .map((row: any[], rowIndex: number) => {
                                  // Check if this is a summary row
                                  const isSummaryRow = row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS');
                                  
                                  return (
                                    <tr 
                                      key={rowIndex} 
                                      className={`${
                                        isSummaryRow 
                                          ? 'bg-blue-50 dark:bg-blue-900/20 font-semibold border-t-2 border-blue-200 dark:border-blue-700' 
                                          : 'hover:bg-gray-50 dark:hover:bg-gray-700'
                                      }`}
                                    >
                                      <td className="px-6 py-4 whitespace-nowrap text-sm">
                                        {!isSummaryRow && (
                                          <input
                                            type="checkbox"
                                            checked={selectedRows.has(`${rowIndex}-9`)}
                                            onChange={(e) => handleRowSelection(`${rowIndex}-9`, e.target.checked)}
                                            className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                          />
                                        )}
                                      </td>
                                      {row.map((cell: any, cellIndex: number) => (
                                        <td
                                          key={cellIndex}
                                          className={`px-6 py-4 whitespace-nowrap text-sm ${
                                            isSummaryRow 
                                              ? 'text-blue-800 dark:text-blue-200 font-medium' 
                                              : 'text-gray-900 dark:text-gray-300'
                                          }`}
                                                                                  >
                                            {cellIndex === 8 ? (
                                            <input
                                              type="text"
                                              value={notes[`${rowIndex}-${cellIndex}`] || ''}
                                              onChange={(e) => handleNoteChange(`${rowIndex}-${cellIndex}`, e.target.value)}
                                              placeholder="Megjegyzés..."
                                              className="w-full min-w-[120px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 9 ? (
                                            <input
                                              type="number"
                                              step="0.1"
                                              value={multipliers[`${rowIndex}-${cellIndex}`] || cell || ''}
                                              onChange={(e) => handleMultiplierChange(`${rowIndex}-${cellIndex}`, Number(e.target.value))}
                                              placeholder="Szorzó"
                                              className="w-full min-w-[80px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 6 || cellIndex === 5 ? formatHungarianDate(cell) : cellIndex === 3 || cellIndex === 4 || cellIndex === 10 ? formatDecimal(cell) : cellIndex === 7 ? formatTime(cell) : cell || ''}
                                            {cellIndex === 10 ? ' Ft' : ''}
                                          </td>
                                      ))}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {activeTable === 'Név szerinti összegzés túlóra kimutatás' && columnsByNameDetailed.length > 0 && (
                        <div className="space-y-4">
                          <h4 className="text-lg font-semibold text-gray-900 dark:text-white">
                            Név szerinti összegzés túlóra kimutatással
                          </h4>
                          <div className="overflow-x-auto">
                            <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700">
                              <thead className="bg-gray-50 dark:bg-gray-700">
                                <tr>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    <input
                                      type="checkbox"
                                      onChange={(e) => handleSelectAll(e.target.checked, 'Név szerinti összegzés túlóra kimutatás')}
                                      className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                    />
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedőkód
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szedő neve
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Dátum
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Túlóra
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                  PDA Pont
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak kezdete
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Műszak vége
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Megjegyzés
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Szorzó
                                  </th>
                                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                                    Fizetés
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-700">
                                {columnsByNameDetailed
                                  .filter((row: any[]) => {
                                    // Filter out completely empty rows
                                    return row && row.some(cell => 
                                      cell !== null && cell !== undefined && cell !== ''
                                    );
                                  })
                                  .map((row: any[], rowIndex: number) => {
                                  // Check if this is a summary row
                                  const isSummaryRow = row[0] && typeof row[0] === 'string' && row[0].includes('ÖSSZEGZÉS');
                                  
                                  return (
                                    <tr 
                                      key={rowIndex} 
                                      className={`${
                                        isSummaryRow 
                                          ? 'bg-blue-50 dark:bg-blue-900/20 font-semibold border-t-2 border-blue-200 dark:border-blue-700' 
                                          : 'hover:bg-gray-50 dark:hover:bg-gray-700'
                                      }`}
                                    >
                                      <td className="px-6 py-4 whitespace-nowrap text-sm">
                                        {!isSummaryRow && (
                                          <input
                                            type="checkbox"
                                            checked={selectedRows.has(`${rowIndex}-8`)}
                                            onChange={(e) => handleRowSelection(`${rowIndex}-8`, e.target.checked)}
                                            className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                          />
                                        )}
                                      </td>
                                      {row.map((cell: any, cellIndex: number) => (
                                        <td
                                          key={cellIndex}
                                          className={`px-6 py-4 whitespace-nowrap text-sm ${
                                            isSummaryRow 
                                              ? 'text-blue-800 dark:text-blue-200 font-medium' 
                                              : 'text-gray-900 dark:text-gray-300'
                                          }`}
                                        >
                                          {cellIndex === 7 ? (
                                            <input
                                              type="text"
                                              value={notes[`${rowIndex}-${cellIndex}`] || ''}
                                              onChange={(e) => handleNoteChange(`${rowIndex}-${cellIndex}`, e.target.value)}
                                              placeholder="Megjegyzés..."
                                              className="w-full min-w-[120px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 8 ? (
                                            <input
                                              type="number"
                                              step="0.1"
                                              value={multipliers[`${rowIndex}-${cellIndex}`] || cell || ''}
                                              onChange={(e) => handleMultiplierChange(`${rowIndex}-${cellIndex}`, Number(e.target.value))}
                                              placeholder="Szorzó"
                                              className="w-full min-w-[80px] px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:text-white"
                                            />
                                          ) : cellIndex === 9 ? formatDecimal(cell) : cellIndex === 3 || cellIndex === 4 ? formatDecimal(cell) : cellIndex === 5 || cellIndex === 6 ? formatHungarianDate(cell, true) : cell || ''}
                                          {cellIndex === 9 ? ' Ft' : ''}
                                          </td>
                                      ))}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
