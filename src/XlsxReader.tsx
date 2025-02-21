import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { collection, addDoc, getDocs, deleteDoc, doc } from 'firebase/firestore';
import { db } from './firebase';

// Define the correct header order as per the XLSX file
const CORRECT_HEADERS = [
  'Scent',
  'Note',
  'Amount',
  'Dupe',
  'Candles',
  'Diffusers',
  'Room Spray',
  'Carpet Freshener',
  'Candle Inventory',
  'timestamp',
];

const XlsxReader: React.FC = () => {
  const [data, setData] = useState<string[][]>([]);
  const [sortConfig, setSortConfig] = useState<{ key: number; direction: 'ascending' | 'descending' } | null>(null);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [searchColumn, setSearchColumn] = useState<number>(0); // Default to the first column
  const [showUpload, setShowUpload] = useState<boolean>(true);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [uploadProgress, setUploadProgress] = useState<number>(0); // Track upload progress
  const [isUploading, setIsUploading] = useState<boolean>(false); // Track if upload is in progress
  const [visibleColumns, setVisibleColumns] = useState<boolean[]>(new Array(CORRECT_HEADERS.length).fill(true)); // Track visible columns

  // Fetch data from Firestore on component mount
  useEffect(() => {
    const fetchData = async () => {
      console.log('Fetching data from Firestore...'); // Debugging log
      try {
        const querySnapshot = await getDocs(collection(db, 'excelData'));
        console.log('Firestore query successful:', querySnapshot); // Debugging log

        const firestoreData: string[][] = [];

        querySnapshot.forEach((doc) => {
          console.log('Document data:', doc.data()); // Debugging log
          const docData = doc.data();
          if (docData) {
            // Map the document data to the correct header order
            const row = CORRECT_HEADERS.map((header) => {
              const value = docData[header];
              if (value && typeof value === 'object' && 'seconds' in value && 'nanoseconds' in value) {
                // Convert Firestore timestamp to a readable date string
                const date = new Date(value.seconds * 1000 + value.nanoseconds / 1000000);
                return date.toLocaleString();
              }
              return String(value || ''); // Use empty string if value is undefined
            });
            firestoreData.push(row);
          }
        });

        console.log('Fetched Firestore Data:', firestoreData); // Debugging log

        if (firestoreData.length > 0) {
          setData(firestoreData);
          setShowUpload(false); // Hide file upload if data exists
        } else {
          setShowUpload(true); // Show file upload if no data exists
        }
      } catch (error) {
        console.error('Error fetching data from Firestore:', error); // Debugging log
      } finally {
        setIsLoading(false); // Data fetching is complete
      }
    };

    fetchData();
  }, []);

  // Function to clear all data from Firestore
  const clearData = async () => {
    // Confirm deletion with the user
    const confirmDelete = window.confirm('Are you sure you want to delete all data? This action cannot be undone.');
    if (!confirmDelete) return;

    try {
      const querySnapshot = await getDocs(collection(db, 'excelData'));
      querySnapshot.forEach(async (doc) => {
        await deleteDoc(doc.ref); // Delete each document
        console.log(`Document with ID ${doc.id} deleted.`);
      });

      // Reset the state
      setData([]);
      setShowUpload(true); // Show the upload input since there's no data
      alert('All data has been deleted successfully.');
    } catch (error) {
      console.error('Error deleting data:', error);
      alert('An error occurred while deleting data.');
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) {
      console.error('No file selected.');
      return;
    }

    const reader = new FileReader();

    reader.onload = async (event: ProgressEvent<FileReader>) => {
      const binaryString = event.target?.result as string;
      console.log('File read successfully:', binaryString); // Debugging log

      try {
        const workbook = XLSX.read(binaryString, { type: 'binary' });
        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];
        const jsonData: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        console.log('Parsed Excel Data:', jsonData); // Debugging log

        // Validate parsed data
        if (jsonData.length === 0) {
          console.error('The file is empty or could not be parsed.');
          return;
        }

        // Check if there are any data rows
        const rows = jsonData.slice(1);
        if (rows.length === 0) {
          console.error('The file has no data rows.');
          return;
        }

        // Convert 2D array to array of objects
        const objectsArray = rows.map((row) => {
          const obj: { [key: string]: string } = {};
          CORRECT_HEADERS.forEach((header, index) => {
            obj[header] = row[index] || ''; // Use empty string if cell is undefined
          });
          return obj;
        });

        console.log('Converted Data:', objectsArray); // Debugging log

        // Start upload process
        setIsUploading(true);
        setUploadProgress(0);

        // Upload each object as a separate document to Firestore
        for (let i = 0; i < objectsArray.length; i++) {
          const obj = objectsArray[i];

          // Skip empty objects
          if (Object.values(obj).every((value) => value === '')) {
            console.warn('Skipping empty row:', obj);
            continue;
          }

          try {
            await addDoc(collection(db, 'excelData'), {
              ...obj,
              timestamp: new Date(),
            });
            console.log(`Document ${i + 1} of ${objectsArray.length} uploaded.`);

            // Update upload progress
            setUploadProgress(((i + 1) / objectsArray.length) * 100);
          } catch (error) {
            console.error('Error adding document: ', error); // Debugging log
          }
        }

        // Update the state to display the new data
        setData(jsonData);
        setShowUpload(false); // Hide file upload after successful upload
      } catch (error) {
        console.error('Error processing or uploading data:', error); // Debugging log
      } finally {
        setIsUploading(false); // Upload process is complete
      }
    };

    reader.onerror = (error) => {
      console.error('Error reading file:', error);
    };

    reader.readAsBinaryString(file);
  };

  const handleSort = (key: number) => {
    let direction: 'ascending' | 'descending' = 'ascending';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };

  const sortedData = React.useMemo(() => {
    if (!sortConfig || !data.length) return data;

    const sorted = [...data].slice(1).sort((a, b) => {
      if (a[sortConfig.key] < b[sortConfig.key]) {
        return sortConfig.direction === 'ascending' ? -1 : 1;
      }
      if (a[sortConfig.key] > b[sortConfig.key]) {
        return sortConfig.direction === 'ascending' ? 1 : -1;
      }
      return 0;
    });

    return [data[0], ...sorted]; // Include the header row
  }, [data, sortConfig]);

  const filteredData = React.useMemo(() => {
    if (!data.length) return []; // Return empty array if no data is loaded

    // Ensure sortedData includes the header row
    const dataToFilter = sortedData.length > 0 ? sortedData : data;

    if (!searchTerm) return dataToFilter;

    return dataToFilter.filter((row, index) => {
      if (index === 0) return true; // Keep the header row
      return row[searchColumn]?.toLowerCase().includes(searchTerm.toLowerCase()); // Filter based on the selected column
    });
  }, [sortedData, searchTerm, searchColumn, data]);

  const toggleColumnVisibility = (index: number) => {
    setVisibleColumns((prevVisibleColumns) => {
      const newVisibleColumns = [...prevVisibleColumns];
      newVisibleColumns[index] = !newVisibleColumns[index];
      return newVisibleColumns;
    });
  };

  // Function to download data as an XLSX file
  const downloadDataAsXLSX = () => {
    if (data.length === 0) {
      alert('No data to download.');
      return;
    }

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Convert the data to a worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Write the workbook to a binary string
    const binaryString = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });

    // Convert the binary string to a Blob
    const blob = new Blob([s2ab(binaryString)], { type: 'application/octet-stream' });

    // Create a download link
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'database.xlsx';
    link.click();

    // Clean up
    URL.revokeObjectURL(url);
  };

  // Utility function to convert a string to an ArrayBuffer
  const s2ab = (s: string) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

  return (
    <div>
      <h1>Upload and Display XLSX File</h1>

      {/* Show loading message while fetching data */}
      {isLoading && <p>Loading data...</p>}

      {/* Show file upload input or table after data is fetched */}
      {!isLoading && (
        <>
          {/* Show file upload input only if there's no data */}
          {showUpload && (
            <div style={{ margin: '20px 0' }}>
              <input type="file" accept=".xlsx" onChange={handleFileUpload} />
            </div>
          )}

          {/* Show "Upload New File", "Clear Data", and "Download Data" buttons if there's data */}
          {!showUpload && (
            <div style={{ margin: '20px 0' }}>
              <button
                onClick={() => setShowUpload(true)}
                style={{ marginRight: '10px', padding: '10px', cursor: 'pointer' }}
              >
                Upload New File
              </button>
              <button
                onClick={clearData}
                style={{ marginRight: '10px', padding: '10px', cursor: 'pointer', backgroundColor: '#ff4444', color: '#fff', border: 'none' }}
              >
                Clear All Data
              </button>
              <button
                onClick={downloadDataAsXLSX}
                style={{ padding: '10px', cursor: 'pointer', backgroundColor: '#4CAF50', color: '#fff', border: 'none' }}
              >
                Download Data as XLSX
              </button>
            </div>
          )}

          {/* Show upload progress bar if upload is in progress */}
          {isUploading && (
            <div style={{ margin: '20px 0', width: '100%', backgroundColor: '#e0e0e0', borderRadius: '4px' }}>
              <div
                style={{
                  width: `${uploadProgress}%`,
                  height: '10px',
                  backgroundColor: '#76c7c0',
                  borderRadius: '4px',
                  transition: 'width 0.3s ease',
                }}
              ></div>
              <p style={{ textAlign: 'center', marginTop: '5px' }}>{Math.round(uploadProgress)}%</p>
            </div>
          )}

          {/* Show search bar and table if there's data */}
          {data.length > 0 && (
            <>
              <div style={{ margin: '20px 0' }}>
                <input
                  type="text"
                  placeholder="Search..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{ padding: '5px', width: '300px', marginRight: '10px' }}
                />
                <select
                  value={searchColumn}
                  onChange={(e) => setSearchColumn(Number(e.target.value))}
                  style={{ padding: '5px' }}
                >
                  {CORRECT_HEADERS.map((header, index) => (
                    <option key={index} value={index}>
                      {header}
                    </option>
                  ))}
                </select>
              </div>
              <div style={{ margin: '20px 0' }}>
                <h3 style={{ color: '#000' }}>Toggle Columns:</h3>
                {CORRECT_HEADERS.map((header, index) => (
                  <label
                    key={index}
                    style={{
                      marginRight: '10px',
                      color: '#000', // Set text color to black
                      display: 'inline-flex', // Align checkbox and label horizontally
                      alignItems: 'center', // Vertically center the checkbox and label
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={visibleColumns[index]}
                      onChange={() => toggleColumnVisibility(index)}
                      style={{ marginRight: '5px' }} // Add spacing between checkbox and label
                    />
                    {header}
                  </label>
                ))}
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      {CORRECT_HEADERS.map((header, index) => (
                        visibleColumns[index] && (
                          <th key={index} onClick={() => handleSort(index)} style={{ cursor: 'pointer', padding: '8px', border: '1px solid #ddd' }}>
                            {header} {sortConfig?.key === index && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                          </th>
                        )
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.slice(1).map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {row.map((cell, cellIndex) => (
                          visibleColumns[cellIndex] && (
                            <td key={cellIndex} style={{ padding: '8px', border: '1px solid #ddd' }}>{cell}</td>
                          )
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
};

export default XlsxReader;