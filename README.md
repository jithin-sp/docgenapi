# Document Generation API

A serverless API that generates Word documents based on form submissions.

## API Endpoint

`POST /api/generate-doc`

## Usage

Send a POST request to the API endpoint with the following JSON structure:

### Contact Form

```json
{
  "formType": "contact",
  "data": {
    "name": "John Doe",
    "email": "john@example.com",
    "message": "Hello, I want to get in touch.",
    "created_at": "2023-05-01T12:00:00Z"
  }
}
```

### Collaboration Form

```json
{
  "formType": "collaboration",
  "data": {
    "name": "Jane Smith",
    "email": "jane@example.com",
    "institution": "University of Example",
    "description": "Research collaboration proposal",
    "created_at": "2023-05-01T12:00:00Z"
  }
}
```

### Concept Form

```json
{
  "formType": "concept",
  "data": {
    "principal_investigators": "Dr. Alice Johnson",
    "email": "alice@example.com",
    "team_members": "Bob, Carol, Dave",
    "project_title": "Novel Research Project",
    "project_description": "This project aims to...",
    "data_required": "Patient records, imaging data",
    "anticipated_outcome": "Development of new treatment methods",
    "created_at": "2023-05-01T12:00:00Z"
  }
}
```

## Response

The API will return a Word document (.docx) file as a downloadable attachment.

## Development

1. Install dependencies:
   ```
   npm install
   ```

2. Test locally using Vercel CLI:
   ```
   vercel dev
   ```

3. Deploy to Vercel:
   ```
   vercel
   ``` 