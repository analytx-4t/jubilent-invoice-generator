// 



import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableCell,
  TableRow,
  WidthType,
  AlignmentType,
  BorderStyle,
} from "docx";
import type { DOCXMergedInvoiceData } from "../types-docx/invoice-docx";

const generateReferenceDebitNoteDOCX = async (
  data: DOCXMergedInvoiceData,
  invoiceType: "main" | "freight",
  formatNumber: (amount: number) => string,
  debitMonth: string
): Promise<Blob> => {
  const isFreight = invoiceType === "freight";
  const amount = isFreight
    ? data.freightBalance
    : (data.loadingCharges || 0) +
      (data.unloadingCharges || 0) +
      (data.localTransportation || 0);
  const subject = isFreight
    ? "Subject: Request for Issuance of Credit Note towards Secondary Freight support/Reimbursement"
    : "Subject: Request for Issuance of Credit Note towards Handling Charges support/Reimbursement";
  const requestText = isFreight
    ? "We request you to kindly issue a credit note towards the support/reimbursement on account of secondary freight expenses incurred by us for the onward movement of goods."
    : "We request you to kindly issue a credit note towards the support/reimbursement on account of Handling Charges incurred by us for the onward movement of goods.";
  const particulars = isFreight
    ? `Secondary Freight Support/Reimbursement for the month ${debitMonth}`
    : `Handling Charges support/reimbursement for the month ${debitMonth}`;
  const text = (value: string, bold = false) =>
    new TextRun({ text: value, bold, font: "Arial", size: 22 });
  const noBorders = {
    top: { style: BorderStyle.NONE },
    bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE },
    right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE },
    insideVertical: { style: BorderStyle.NONE },
  };
  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [text(data.customer.customerName, true)], spacing: { after: 160 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [text(`ADDRESS: ${data.customer.address || ""}`, true)], spacing: { after: 100 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [text(`GSTIN: ${data.customer.gstin}  PAN: ${data.customer.pan}  MOB: ${data.customer.mobile}`, true)],
          spacing: { after: 260 },
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [new TableRow({ children: [
            new TableCell({ children: [new Paragraph({ text: "To," })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [text("Dated:", true)] })] }),
          ] })],
        }),
        ...[
          "The Sales Head",
          "Jubilant Agri & Consumer Products Ltd.",
          "Plot No 142, Chimes, 3rd Floor, Sector 44,",
          "Gurugram Haryana-122003",
        ].map((line) => new Paragraph({ text: line })),
        new Paragraph({ text: subject, spacing: { before: 320, after: 260 } }),
        new Paragraph({ text: "Dear sir/madam,", spacing: { after: 240 } }),
        new Paragraph({ text: requestText, spacing: { after: 300 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [text("Particulars", true)] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [text("Amt. INR", true)] })] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ text: particulars })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, text: formatNumber(amount) })] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [text("Total", true)] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [text(formatNumber(amount), true)] })] }),
            ] }),
          ],
        }),
        new Paragraph({ text: "Thanking You", spacing: { before: 500 } }),
        new Paragraph({ text: "Yours Faithfully" }),
        new Paragraph({ text: "Authorised Signatory", spacing: { before: 300 } }),
      ],
    }],
  });
  return Packer.toBlob(doc);
};

export const generateDebitNoteDOCX = async (
  data: DOCXMergedInvoiceData,
  invoiceType: "godown" | "main" | "freight",
  formatNumber: (amount: number) => string,
  debitMonth = ""
): Promise<Blob> => {
  if ((invoiceType as string) === "main" || (invoiceType as string) === "freight") {
      return generateReferenceDebitNoteDOCX(
        data,
        invoiceType as "main" | "freight",
        formatNumber,
        debitMonth
      );
  }

  // Helper: create cell with consistent styling
  const createCell = (
    text: string,
    options: {
      bold?: boolean;
      fontSize?: number; // in half-points (20 = 10pt)
      alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
      width?: number; // in DXA (1440 DXA = 1 inch)
    } = {}
  ): TableCell => {
    const { bold = false, fontSize = 20, alignment = AlignmentType.LEFT, width } = options;
    return new TableCell({
      width: width ? { size: width, type: WidthType.DXA } : undefined,
      children: [
        new Paragraph({
          alignment,
          children: [
            new TextRun({
              text,
              bold,
              font: "Arial",
              size: fontSize,
            }),
          ],
        }),
      ],
    });
  };

  // Fixed Jubilant address (same as PDF)
  const fixedJubilantAddress = [
    "The Sales Head",
    "Jubilant Agri & Consumer Products Ltd.",
    "Plot No 142, Chimes, 3rd Floor, Sector 44,",
    "Gurugram 3rd Floor, Sector -44, Gurugram,",
    "Haryana-122003",
  ];

  // Month hardcoded as in PDF (you can make dynamic if needed)
  const monthYear = debitMonth;

  // Build introductory text with month
  let introductoryText = "";
  if (invoiceType === "godown") {
    introductoryText = `Request to kindly reimburse the Godown Rent Expenses for the month of ${monthYear} as per details given as under:`;
  } else if (invoiceType === "main") {
    introductoryText = `Request to kindly reimburse the Handling Expenses for the month of ${monthYear} as per details given as under:`;
  } else if (invoiceType === "freight") {
    introductoryText = `Request to kindly reimburse the Freight Expenses for the month of ${monthYear} as per details given as below:`;
  }

  // Build subject
  let subjectText = "";
  if (invoiceType === "godown") {
    subjectText = "Subject: Reimbursement of Rent Expenses";
  } else if (invoiceType === "main") {
    subjectText = "Subject: Reimbursement of Handling Expenses.";
  } else if (invoiceType === "freight") {
    subjectText = "Subject: Reimbursement of Freight.";
  }

  // Build table based on invoice type
  let tableHeaders: TableCell[] = [];
  let serviceRows: TableRow[] = [];
  let totalAmountBeforeTax = 0;

  if (invoiceType === "freight") {
    // Only Particulars + Amount
    tableHeaders = [
      createCell("Particulars", { bold: true, fontSize: 20, width: 9000 }),
      createCell("Amt. In Ru.", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 3000,
      }),
    ];

    serviceRows.push(
      new TableRow({
        children: [
          createCell("REIMBURSEMENT OF FREIGHT EXPENSES", {
            fontSize: 20,
            width: 9000,
          }),
          createCell(formatNumber(data.freightBalance), {
            fontSize: 20,
            alignment: AlignmentType.RIGHT,
            width: 3000,
          }),
        ],
      })
    );
    totalAmountBeforeTax = data.freightBalance;
  } else if (invoiceType === "main") {
    // Particulars + Qty + Amount (no Rate)
    tableHeaders = [
      createCell("Particulars", { bold: true, fontSize: 20, width: 8000 }),
      createCell("Qty in MT", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.CENTER,
        width: 2000,
      }),
      createCell("Amt. In Ru.", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 2000,
      }),
    ];

    const totalHandling =
      (data.loadingCharges || 0) +
      (data.unloadingCharges || 0) +
      (data.localTransportation || 0);

    // const quantity = formatNumber(
    //   (data.loadingCharges || 0) / 75 +
    //     (data.unloadingCharges || 0) / 75 +
    //     (data.localTransportation || 0) / 200
    // );
    const quantity = formatNumber(data.quantityLifted || 0);

    serviceRows.push(
      new TableRow({
        children: [
          createCell(
            "REIMBURSEMENT OF LOADING, UNLOADING & LOCAL TRANSPORTATION",
            { fontSize: 20, width: 8000 }
          ),
          createCell(quantity, {
            fontSize: 20,
            alignment: AlignmentType.CENTER,
            width: 2000,
          }),
          createCell(formatNumber(totalHandling), {
            fontSize: 20,
            alignment: AlignmentType.RIGHT,
            width: 2000,
          }),
        ],
      })
    );
    totalAmountBeforeTax = totalHandling;
  } else {
    // Godown: all 4 columns
    tableHeaders = [
      createCell("Particulars", { bold: true, fontSize: 20, width: 7000 }),
      createCell("Qty in MT", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.CENTER,
        width: 2000,
      }),
      createCell("Rate/MT", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.CENTER,
        width: 2000,
      }),
      createCell("Amt. In Ru.", {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 2000,
      }),
    ];

    serviceRows.push(
      new TableRow({
        children: [
          createCell("REIMBURSEMENT OF RENT EXPENSES", {
            fontSize: 20,
            width: 7000,
          }),
          createCell(formatNumber(data.quantityLifted), {
            fontSize: 20,
            alignment: AlignmentType.CENTER,
            width: 2000,
          }),
          createCell("100.00", {
            fontSize: 20,
            alignment: AlignmentType.CENTER,
            width: 2000,
          }),
          createCell(formatNumber(data.godownRent), {
            fontSize: 20,
            alignment: AlignmentType.RIGHT,
            width: 2000,
          }),
        ],
      })
    );
    totalAmountBeforeTax = data.godownRent;
  }

  // Total row (adapt to column count)
  const totalCells: TableCell[] = [];
  if (invoiceType === "freight") {
    totalCells.push(
      createCell("Total", { bold: true, fontSize: 20, width: 9000 }),
      createCell(formatNumber(totalAmountBeforeTax), {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 3000,
      })
    );
  } else if (invoiceType === "main") {
    totalCells.push(
      createCell("Total", { bold: true, fontSize: 20, width: 8000 }),
      createCell("", { width: 2000 }), // empty qty
      createCell(formatNumber(totalAmountBeforeTax), {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 2000,
      })
    );
  } else {
    totalCells.push(
      createCell("Total", { bold: true, fontSize: 20, width: 7000 }),
      createCell("", { width: 2000 }),
      createCell("", { width: 2000 }),
      createCell(formatNumber(totalAmountBeforeTax), {
        bold: true,
        fontSize: 20,
        alignment: AlignmentType.RIGHT,
        width: 2000,
      })
    );
  }

  const children = [
    // Customer Name - Centered
    new Paragraph({
       alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [
    new TextRun({
      text: data.customer.customerName,
      bold: true,
      size: 24, // optional: matches Heading1 ~12pt
      font: "Arial",
      color: "000000", // black
    }),
  ]
    }),

    // Customer Address - Centered
    new Paragraph({
       alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [
    new TextRun({
      text: `ADDRESS: ${data.customer.address}`,
      bold: true,
      size: 20,
      font: "Arial",
      color: "000000", // black
    }),
  ],
    }),

    // GSTIN/PAN/MOB - Centered
    new Paragraph({
       alignment: AlignmentType.CENTER,
  spacing: { after: 300 },
  children: [
    new TextRun({
      text: `GSTIN: ${data.customer.gstin}   PAN: ${data.customer.pan}${
        data.customer.mobile ? `   MOB: ${data.customer.mobile}` : ""
      }`,
      bold: true,
      size: 20,
      font: "Arial",
      color: "000000", // black
    }),
  ],
    }),

    // "To:" Section - Left aligned
    new Paragraph({
      text: "To:",
      spacing: { after: 100 },
    }),

    // Fixed Jubilant Address (no GSTIN)
    ...fixedJubilantAddress.map(
      (line) =>
        new Paragraph({
          text: line,
          spacing: { after: 100 },
        })
    ),

    // Subject
    new Paragraph({
      text: subjectText,
      spacing: { after: 300 },
    }),

    // Dear Sir
    new Paragraph({
      text: "Dear Sir",
      spacing: { after: 200 },
    }),

    // Introductory text
    new Paragraph({
      text: introductoryText,
      spacing: { after: 300 },
    }),

    // Table
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
      rows: [
        new TableRow({ children: tableHeaders }),
        ...serviceRows,
        new TableRow({ children: totalCells }),
      ],
    }),

    // Footer
    new Paragraph({ text: "Thanking You", spacing: { after: 200 } }),
    new Paragraph({ text: "Yours Faithfully", spacing: { after: 200 } }),
    new Paragraph({ text: "Authorised Signatory", spacing: { after: 400 } }),
  ];

  const doc = new Document({
    sections: [{ children }],
  });

  return await Packer.toBlob(doc);
};