# goExcel: Excel File Handling API

## Overview

goExcel is a Go-based RESTful API that facilitates the upload, extraction, and processing of data from Excel files with unknown structures. This project is designed to handle Excel files (.xlsx) with varying column counts and types, making it versatile for different data import scenarios.

## Features

- File upload endpoint for Excel (.xlsx) files
- Automatic extraction of data from all sheets in the uploaded Excel file
- Handling of unknown Excel structures with variable column counts and types
- Basic processing and structuring of extracted data
- JSON response with processed data from all sheets

## Prerequisites

- Go (version 1.16 or later)
- Git

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/goExcel.git
   cd goExcel
   ```

2. Install dependencies:
   ```
   go mod tidy
   ```

## Running the Application

To start the server, run:

```
go run main.go
```

The server will start in debug mode and listen on `http://localhost:8080`.

## API Endpoints

### Upload Excel File

- **URL**: `/upload`
- **Method**: `POST`
- **Content-Type**: `multipart/form-data`
- **Parameter**: 
  - `file`: The Excel file to upload (must be .xlsx format)

#### Success Response

- **Code**: 200 OK
- **Content**: JSON object containing the processed data from all sheets

```json
{
  "message": "File processed successfully",
  "data": [
    {
      "sheetName": "Sheet1",
      "headers": ["Column1", "Column2", ...],
      "data": [
        {"Column1": "Value1", "Column2": "Value2", ...},
        ...
      ]
    },
    ...
  ]
}
```

#### Error Responses

- **Code**: 400 Bad Request
  - **Content**: `{"error": "No file uploaded"}`
  - **Content**: `{"error": "Invalid file format. Please upload an Excel file (.xlsx)"}`

- **Code**: 500 Internal Server Error
  - **Content**: `{"error": "Failed to save file"}`
  - **Content**: `{"error": "Failed to open Excel file"}`
  - **Content**: `{"error": "Error processing sheet [sheet_name]: [error_message]"}`

## Testing

### Using REST Client (VS Code Extension)

1. Install the REST Client extension in VS Code.
2. Create a new file named `test.http` or `test.rest`.
3. Add the following content to the file:

```http
### Upload Excel File
POST http://localhost:8080/upload
Content-Type: multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW

------WebKitFormBoundary7MA4YWxkTrZu0gW
Content-Disposition: form-data; name="file"; filename="test.xlsx"
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

< ./path/to/your/test.xlsx
------WebKitFormBoundary7MA4YWxkTrZu0gW--
```

4. Replace `./path/to/your/test.xlsx` with the actual path to your Excel file.
5. Click on "Send Request" above the POST line to send the request.

### Using curl

You can also test the endpoint using curl:

```bash
curl -X POST -F "file=@path/to/your/test.xlsx" http://localhost:8080/upload
```

Replace `path/to/your/test.xlsx` with the actual path to your Excel file.

## Error Handling

The application includes error handling for various scenarios, including:
- Missing file in the upload request
- Invalid file format (non-Excel files)
- File saving errors
- Excel file processing errors

Detailed error messages are returned in the API responses and logged in the console.

## Limitations

- The application currently supports only .xlsx file formats.
- Large Excel files may take longer to process and may require additional memory.

## Future Improvements

- Add support for other Excel formats (.xls, .csv)
- Implement concurrent processing for large files
- Add more advanced data processing options
- Implement user authentication and file access controls
