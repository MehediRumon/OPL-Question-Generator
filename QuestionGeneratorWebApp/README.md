# Question Generator Web Application

A modern web application for generating MCQ (Multiple Choice Questions) and SAQ (Short Answer Questions) documents using Microsoft Word templates.

## Features

- **Generate MCQ Documents**: Create multiple choice question documents with customizable subjects and sequences
- **Generate SAQ Documents**: Create short answer question documents with Bangla and English support
- **Add Answer Tags**: Automatically add answer tags to existing Word documents
- **Add Subject Names**: Add subject names with serial numbers to question documents
- **Web-based Interface**: Easy-to-use web interface with modern design
- **REST API**: Full REST API for programmatic access

## Requirements

- .NET 8.0 SDK or later
- **Microsoft Word** (must be installed on the system)
- **Windows environment** (Office Interop only works on Windows)
- **Office Primary Interop Assemblies** (typically installed with Office)

## Getting Started

### Installation

1. Clone the repository
2. Navigate to the QuestionGeneratorWebApp directory
3. **Ensure Microsoft Office is installed on your system**
4. Restore dependencies:
   ```bash
   dotnet restore
   ```

### Running the Application

1. Clean and build the application (important after first clone):
   ```bash
   dotnet clean
   dotnet build
   ```

2. Run the application:
   ```bash
   dotnet run
   ```

3. Open your browser and navigate to:
   - Web UI: `http://localhost:5000` or `https://localhost:5001`
   - Swagger API: `http://localhost:5000/swagger` or `https://localhost:5001/swagger`

### Troubleshooting

If you encounter errors like "Could not load file or assembly 'office'", please refer to the comprehensive [TROUBLESHOOTING.md](TROUBLESHOOTING.md) guide which covers:
- Installing Office Primary Interop Assemblies
- DCOM configuration
- Permission issues
- Architecture mismatches
- Server deployment specific issues

## Usage

### Web Interface

The web interface provides four main tabs:

1. **Generate MCQ**: Create MCQ question documents
   - Specify number of questions
   - Define subject list (e.g., "phy, chem, math, bio")
   - Set sequence ranges (e.g., "1-25,26-50,51-75,76-100")

2. **Generate SAQ**: Create SAQ question documents
   - Specify number of questions and marks
   - Define Bangla and English subject lists
   - Set sequence ranges

3. **Add Answer Tag**: Upload a document to add answer tags
   - Upload a .docx file
   - System automatically identifies answers and adds tags

4. **Add Subject Name**: Upload a document to add subject names
   - Upload a .docx file
   - Specify subject list and sequences
   - System adds subject names to appropriate tables

### API Endpoints

#### Generate MCQ
```
POST /api/QuestionGenerator/generate-mcq
Content-Type: application/json

{
  "questionNumber": 100,
  "subjectList": "phy, chem, math, bio",
  "sequenceList": "1-25,26-50,51-75,76-100"
}
```

#### Generate SAQ
```
POST /api/QuestionGenerator/generate-saq
Content-Type: application/json

{
  "questionNumber": 100,
  "questionMark": 2,
  "subjectListBangla": "পদার্থবিজ্ঞান, রসায়ন",
  "subjectListEnglish": "Phy,Chem",
  "sequenceList": "1-50,51-100"
}
```

#### Add Answer Tag
```
POST /api/QuestionGenerator/add-answer-tag
Content-Type: multipart/form-data

file: [Word document]
```

#### Add Subject Name
```
POST /api/QuestionGenerator/add-subject-name
Content-Type: multipart/form-data

file: [Word document]
subjectList: "phy, chem, math, bio"
sequenceList: "1-25,26-50,51-75,76-100"
```

#### Download File
```
GET /api/QuestionGenerator/download/{fileName}
```

## Project Structure

```
QuestionGeneratorWebApp/
├── Controllers/
│   └── QuestionGeneratorController.cs  # API endpoints
├── Services/
│   ├── McqQuestionGenerator.cs         # MCQ generation logic
│   └── SaqQuestionGenerator.cs         # SAQ generation logic
├── Question/
│   ├── McqSample.docx                  # MCQ template
│   └── SaqSample.docx                  # SAQ template
├── wwwroot/
│   └── index.html                      # Web UI
├── Generated/                           # Output directory for generated files
├── Uploads/                            # Upload directory for processed files
└── Program.cs                          # Application entry point
```

## Configuration

The application uses the following default settings:

- **MCQ Defaults**:
  - Question Number: 100
  - Subject List: "phy, chem, math, bio"
  - Sequence List: "1-25,26-50,51-75,76-100"

- **SAQ Defaults**:
  - Question Number: 100
  - Question Mark: 2
  - Subject List (Bangla): "পদার্থবিজ্ঞান, রসায়ন"
  - Subject List (English): "Phy,Chem"
  - Sequence List: "1-50,51-100"

## Technical Details

- **Framework**: ASP.NET Core 8.0
- **API**: REST API with Swagger documentation
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
- **Document Processing**: Microsoft.Office.Interop.Word (with `EmbedInteropTypes=False` for web deployment)
- **File Management**: Multi-part form data for uploads

## Notes

- Generated documents are saved in the `Generated` folder
- Uploaded documents are processed and saved in the `Uploads` folder
- Sample templates (McqSample.docx and SaqSample.docx) must be present in the `Question` folder
- The application requires Microsoft Word to be installed for Office Interop functionality
- **Important**: The project is configured with `EmbedInteropTypes=False` for the Office Interop package to ensure proper assembly loading in web deployments

## License

This project is provided as-is for educational and professional use.
