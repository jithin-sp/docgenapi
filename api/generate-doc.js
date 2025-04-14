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
    
    try {
      // Use Packer to generate a buffer with compatibility mode
      const buffer = await Packer.toBuffer(doc, {
        // Add specific configuration for packing that improves Google Docs compatibility
        compatibility: true
      });

      // Set response headers for file download with caching disabled to prevent preview issues
      res.setHeader('Content-Disposition', `attachment; filename="${formType}-form.docx"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
      
      // Send the document
      res.status(200).send(buffer);
    } catch (packingError) {
      console.error('Error packing document:', packingError);
      res.status(500).json({ error: 'Failed to package document' });
    }

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
    const leftLogoWidth = 1.5 * 72; // Reduced size for better compatibility
    const leftLogoHeight = 0.4 * 72; // Reduced size for better compatibility
    const rightLogoWidth = 1.5 * 72; // Reduced size for better compatibility
    const rightLogoHeight = 0.4 * 72; // Reduced size for better compatibility
    
    try {
      const leftBuffer = fs.readFileSync(logoLeftPath);
      const rightBuffer = fs.readFileSync(logoRightPath);
      
      // Create visible header with both logos using a table with visible borders
      const headerTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
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
                  top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                  right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
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
        top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              },
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
              borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              },
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
  
  // Add data rows based on form type in a two-column layout for better compatibility
  Object.entries(data).forEach(([key, value]) => {
    rows.push(
      new TableRow({
        children: [
          // First column: Field name
          new TableCell({
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
            },
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: {
              fill: "F0F0F0", // Light gray background
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: key.replace(/_/g, " ").replace(/\b\w/g, l => l.toUpperCase()),
                    bold: true,
                    size: 24,
                    font: "Calibri",
                  }),
                ],
              }),
            ],
          }),
          // Second column: Value
          new TableCell({
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
            },
            width: { size: 70, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: value ? value.toString() : "",
                    size: 24,
                    font: "Calibri",
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
      top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
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
          size: 18, // Slightly smaller font for better compatibility
          font: "Arial", // More standard font for better compatibility
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
        size: 32, // Slightly smaller for better compatibility
        font: "Arial", // More standard font for better compatibility
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
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
    compatibility: { doNotUseFloatingTableGrid: true }, // Improves compatibility with Google Docs
    externalStyles: false, // Simplifies style management
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1000,
            right: 1000,
            bottom: 1000,
            left: 1000,
          },
        },
      },
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
            size: 22, // Slightly smaller font for better compatibility
            font: "Arial", // More standard font for better compatibility
          },
          paragraph: {
            spacing: {
              line: 276, // Reduced line spacing for better compatibility
            },
          },
        },
      ],
    },
  });

  // Return the document (not a buffer)
  return doc;
} 