const { Document, Paragraph, TextRun, Packer, Table, TableRow, TableCell, BorderStyle, WidthType, AlignmentType, ImageRun } = require('docx');
const fs = require('fs');
const path = require('path');

// Handler for serverless function
module.exports = async (req, res) => {
  try {
    // Set CORS headers
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,POST');
    res.setHeader(
      'Access-Control-Allow-Headers',
      'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
    );

    // Handle preflight request
    if (req.method === 'OPTIONS') {
      res.status(200).end();
      return;
    }

    // Ensure the request is POST
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed. Please use POST.' });
    }

    // Get form data from request body
    const { formType, data } = req.body;

    if (!formType || !data) {
      return res.status(400).json({ error: 'Missing required fields: formType or data' });
    }

    // Generate the document based on form type
    const doc = await generateDocument(formType, data);
    
    // Use Packer to generate a buffer
    const buffer = await Packer.toBuffer(doc);

    // Set response headers for file download
    res.setHeader('Content-Disposition', `attachment; filename="${formType}-form.docx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    
    // Send the document
    res.send(buffer);

  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document' });
  }
};

// Function to create header with logos
function createHeaderWithLogos() {
  try {
    const logoLeftPath = path.resolve('./logos/logo_left.png');
    const logoRightPath = path.resolve('./logos/logo_right.png');
    
    // Make sure the image files exist
    if (!fs.existsSync(logoLeftPath)) {
      console.warn('Left logo file not found:', logoLeftPath);
      return createLogoPlaceholders();
    }
    if (!fs.existsSync(logoRightPath)) {
      console.warn('Right logo file not found:', logoRightPath);
      return createLogoPlaceholders();
    }
    
    // Convert inches to points (1 inch = 72 points)
    const leftLogoWidth = 3.14 * 72; // 226.08 points
    const leftLogoHeight = 0.67 * 72; // 48.24 points
    const rightLogoWidth = 2.86 * 72; // 205.92 points
    const rightLogoHeight = 0.63 * 72; // 45.36 points
    
    try {
      const leftBuffer = fs.readFileSync(logoLeftPath);
      const rightBuffer = fs.readFileSync(logoRightPath);
      
      // Create visible header with both logos using a table
      const headerTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
          insideHorizontal: { style: BorderStyle.NONE },
          insideVertical: { style: BorderStyle.NONE },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                children: [
                  new Paragraph({
                    children: [
                      new ImageRun({
                        data: leftBuffer,
                        transformation: {
                          width: leftLogoWidth,
                          height: leftLogoHeight,
                        },
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                children: [
                  new Paragraph({
                    children: [
                      new ImageRun({
                        data: rightBuffer,
                        transformation: {
                          width: rightLogoWidth,
                          height: rightLogoHeight,
                        },
                      }),
                    ],
                    alignment: AlignmentType.RIGHT,
                  }),
                ],
              }),
            ],
          }),
        ],
      });
      
      return [
        headerTable,
        new Paragraph({ text: "" }),
      ];
    } catch (error) {
      console.error('Error embedding logo images in document:', error);
      return createLogoPlaceholders();
    }
  } catch (error) {
    console.error('Error in header creation:', error);
    return createLogoPlaceholders();
  }
}

// Create text placeholders for logos
function createLogoPlaceholders() {
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "[LEFT LOGO]",
                      bold: true,
                      size: 24,
                    }),
                  ],
                  alignment: AlignmentType.LEFT,
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "[RIGHT LOGO]",
                      bold: true,
                      size: 24,
                    }),
                  ],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ text: "" }),
  ];
}

// Function to create a data table
function createDataTable(data, formType) {
  const rows = [];
  
  // Add data rows based on form type in a single column layout (field name, then value)
  Object.entries(data).forEach(([key, value]) => {
    // Field name row with light gray background
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
            },
            width: { size: 100, type: WidthType.PERCENTAGE },
            shading: {
              fill: "F0F0F0", // Light gray background
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: key.replace(/_/g, " ").replace(/\b\w/g, l => l.toUpperCase()),
                    bold: true,
                    size: 24, // Bigger font
                    font: "Calibri", // More readable font
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    );
    
    // Value row (white background)
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
            },
            width: { size: 100, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: value.toString(),
                    size: 24, // Bigger font
                    font: "Calibri", // More readable font
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    );
  });
  
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "F0F0F0" }, // Very light gray or remove completely
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: rows,
  });
}

// Function to create confidentiality notice
function createConfidentialityNotice() {
  return [
    new Paragraph({ text: "" }), // Spacer
    new Paragraph({ text: "" }), // Spacer
    new Paragraph({
      children: [
        new TextRun({
          text: "Confidential: This document contains sensitive information. Please share only with trusted parties and with caution.",
          italics: true,
          color: "808080", // Gray text
          size: 20, // Bigger font
          font: "Calibri", // More readable font
        }),
      ],
      alignment: AlignmentType.CENTER,
    }),
  ];
}

// Function to generate documents based on form type
async function generateDocument(formType, data) {
  // Calculate current date if not provided in data
  if (formType === 'collaboration' && !data.created_at) {
    data.created_at = new Date().toISOString();
  }
  
  // Create document title based on form type
  const documentTitle = new Paragraph({
    children: [
      new TextRun({
        text: `${formType.toUpperCase()} FORM`,
        bold: true,
        size: 36, // Bigger font
        font: "Calibri", // More readable font
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
  });
  
  // Create document sections
  let children = [
    ...createHeaderWithLogos(),
    documentTitle,
    createDataTable(data, formType),
    ...createConfidentialityNotice(),
  ];
  
  const doc = new Document({
    creator: "CDHPM DocGen",
    description: "Automatically generated document",
    title: `${formType} Form`,
    sections: [{
      properties: {},
      children: children
    }],
    styles: {
      paragraphStyles: [
        {
          id: "Normal",
          name: "Normal",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            size: 24, // Default size for better readability
            font: "Calibri",
          },
          paragraph: {
            spacing: {
              line: 360, // Improved line spacing
            },
          },
        },
      ],
    },
  });

  // Return the document (not a buffer)
  return doc;
} 