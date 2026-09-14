// // // // import {
// // // //   Document,
// // // //   Packer,
// // // //   Paragraph,
// // // //   TextRun,
// // // //   Table,
// // // //   TableCell,
// // // //   TableRow,
// // // //   WidthType,
// // // //   VerticalAlign,
// // // //   AlignmentType,
// // // //   BorderStyle,
// // // //   HeightRule,
// // // // } from "docx";
// // // // import type { DOCXMergedInvoiceData } from "../types-docx/invoice-docx";

// // // // export const generateInvoiceDOCX = async (
// // // //   data: DOCXMergedInvoiceData,
// // // //   invoiceType: "godown" | "main" | "freight",
// // // //   formatNumber: (amount: number) => string
// // // // ): Promise<Blob> => {
// // // //   console.log("🎯 Starting DOCX generation for:", invoiceType);
// // // //   console.log("📊 Data received:", {
// // // //     godownRent: data.godownRent,
// // // //     loadingCharges: data.loadingCharges,
// // // //     unloadingCharges: data.unloadingCharges,
// // // //     localTransportation: data.localTransportation,
// // // //     freightBalance: data.freightBalance,
// // // //     customer: data.customer.customerName,
// // // //   });
// // // //   // Helper function to create table cell with text
// // // //   const createCell = (
// // // //     text: string,
// // // //     options: {
// // // //       bold?: boolean;
// // // //       fontSize?: number;
// // // //       alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
// // // //       width?: number;
// // // //       colSpan?: number;
// // // //     } = {}
// // // //   ): TableCell => {
// // // //     const {
// // // //       bold = false,
// // // //       fontSize = 16,
// // // //       alignment = AlignmentType.LEFT,
// // // //       width,
// // // //       colSpan,
// // // //     } = options;

// // // //     return new TableCell({
// // // //       width: width ? { size: width, type: WidthType.DXA } : undefined,
// // // //       columnSpan: colSpan,
// // // //       children: [
// // // //         new Paragraph({
// // // //           alignment,
// // // //           children: [
// // // //             new TextRun({
// // // //               text,
// // // //               bold,
// // // //               font: "Arial",
// // // //               size: fontSize * 2, // docx uses half-points
// // // //             }),
// // // //           ],
// // // //         }),
// // // //       ],
// // // //     });
// // // //   };

// // // //   // Helper function to create empty cell for spacing
// // // //   const createEmptyCell = (width?: number): TableCell => {
// // // //     return new TableCell({
// // // //       width: width ? { size: width, type: WidthType.DXA } : undefined,
// // // //       children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
// // // //     });
// // // //   };

// // // //   const customerState = data.customer.address.includes("Maharashtra")
// // // //     ? "Maharashtra"
// // // //     : data.customer.address.includes("Gujarat")
// // // //     ? "Gujarat"
// // // //     : data.customer.address.includes("Chhattisgarh")
// // // //     ? "Chhattisgarh"
// // // //     : data.customer.address.includes("Uttarakhand")
// // // //     ? "Uttarakhand"
// // // //     : "Uttar Pradesh";

// // // //   // Calculate amounts
// // // //   let serviceRows: TableRow[] = [];
// // // //   let totalAmountBeforeTax = 0;

// // // //   switch (invoiceType) {
// // // //     case "godown":
// // // //       serviceRows.push(
// // // //         new TableRow({
// // // //           children: [
// // // //             createCell(
// // // //               "Rental or Leasing services involving own or leased non - residential property",
// // // //               {
// // // //                 bold: false,
// // // //                 fontSize: 14,
// // // //                 width: 6000,
// // // //               }
// // // //             ),
// // // //             createCell("997212", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber((data.godownRent || 0) / 100), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell("100.00", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber(data.godownRent), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //           ],
// // // //         })
// // // //       );
// // // //       totalAmountBeforeTax = data.godownRent;
// // // //       break;

// // // //     case "main":
// // // //       serviceRows.push(
// // // //         new TableRow({
// // // //           children: [
// // // //             createCell("Loading Charges", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               width: 6000,
// // // //             }),
// // // //             createCell("996519", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber((data.loadingCharges || 0) / 75), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell("75.00", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber(data.loadingCharges), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //           ],
// // // //         }),
// // // //         new TableRow({
// // // //           children: [
// // // //             createCell("Unloading Charges", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               width: 6000,
// // // //             }),
// // // //             createCell("996519", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber((data.unloadingCharges || 0) / 75), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell("75.00", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber(data.unloadingCharges), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //           ],
// // // //         }),
// // // //         new TableRow({
// // // //           children: [
// // // //             createCell("Local Transportation", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               width: 6000,
// // // //             }),
// // // //             createCell("996713", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber((data.localTransportation || 0) / 200), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell("200.00", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber(data.localTransportation), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 1500,
// // // //             }),
// // // //           ],
// // // //         })
// // // //       );
// // // //       totalAmountBeforeTax =
// // // //         data.loadingCharges + data.unloadingCharges + data.localTransportation;
// // // //       break;

// // // //     case "freight":
// // // //       serviceRows.push(
// // // //         new TableRow({
// // // //           children: [
// // // //             createCell("Secondary Freight", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               width: 6000,
// // // //             }),
// // // //             createCell("996511", {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.CENTER,
// // // //               width: 1500,
// // // //             }),
// // // //             createCell(formatNumber(data.freightBalance), {
// // // //               bold: false,
// // // //               fontSize: 14,
// // // //               alignment: AlignmentType.RIGHT,
// // // //               width: 3000,
// // // //               colSpan: 2,
// // // //             }),
// // // //           ],
// // // //         })
// // // //       );
// // // //       totalAmountBeforeTax = data.freightBalance;
// // // //       break;
// // // //   }

// // // //   // Calculate taxes
// // // //   const interStateStates = [
// // // //     "Maharashtra",
// // // //     "Gujarat",
// // // //     "Chhattisgarh",
// // // //     "Uttarakhand",
// // // //   ];
// // // //   const isInterState = interStateStates.includes(customerState);

// // // //   let cgst = 0;
// // // //   let sgst = 0;
// // // //   let igst = 0;
// // // //   let totalAmount = 0;

// // // //   if (isInterState) {
// // // //     igst = totalAmountBeforeTax * 0.18;
// // // //     totalAmount = totalAmountBeforeTax + igst;
// // // //   } else {
// // // //     cgst = totalAmountBeforeTax * 0.09;
// // // //     sgst = totalAmountBeforeTax * 0.09;
// // // //     totalAmount = totalAmountBeforeTax + cgst + sgst;
// // // //   }

// // // //   // Create tax rows
// // // //   const taxRows: TableRow[] = [];

// // // //   if (isInterState) {
// // // //     taxRows.push(
// // // //       new TableRow({
// // // //         children: [
// // // //           createCell("IGST @ 18%", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             width: 7500,
// // // //             colSpan: 4,
// // // //           }),
// // // //           createCell(formatNumber(igst), {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.RIGHT,
// // // //             width: 1500,
// // // //           }),
// // // //         ],
// // // //       })
// // // //     );
// // // //   } else {
// // // //     taxRows.push(
// // // //       new TableRow({
// // // //         children: [
// // // //           createCell("CGST @ 9%", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             width: 7500,
// // // //             colSpan: 4,
// // // //           }),
// // // //           createCell(formatNumber(cgst), {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.RIGHT,
// // // //             width: 1500,
// // // //           }),
// // // //         ],
// // // //       }),
// // // //       new TableRow({
// // // //         children: [
// // // //           createCell("SGST @ 9%", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             width: 7500,
// // // //             colSpan: 4,
// // // //           }),
// // // //           createCell(formatNumber(sgst), {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.RIGHT,
// // // //             width: 1500,
// // // //           }),
// // // //         ],
// // // //       })
// // // //     );
// // // //   }

// // // //   // Create main table
// // // //   const table = new Table({
// // // //     width: {
// // // //       size: 100,
// // // //       type: WidthType.PERCENTAGE,
// // // //     },
// // // //     borders: {
// // // //       top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //       bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //       left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //       right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //       insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //       insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
// // // //     },
// // // //     rows: [
// // // //       // Header row
// // // //       new TableRow({
// // // //         children: [
// // // //           createCell("Service Description", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             width: 6000,
// // // //           }),
// // // //           createCell("HSN / SAC", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.CENTER,
// // // //             width: 1500,
// // // //           }),
// // // //           ...(invoiceType === "freight"
// // // //             ? [
// // // //                 createCell("Amount", {
// // // //                   bold: true,
// // // //                   fontSize: 14,
// // // //                   alignment: AlignmentType.CENTER,
// // // //                   width: 3000,
// // // //                   colSpan: 2,
// // // //                 }),
// // // //               ]
// // // //             : [
// // // //                 createCell("Qty", {
// // // //                   bold: true,
// // // //                   fontSize: 14,
// // // //                   alignment: AlignmentType.CENTER,
// // // //                   width: 1500,
// // // //                 }),
// // // //                 createCell("Rate", {
// // // //                   bold: true,
// // // //                   fontSize: 14,
// // // //                   alignment: AlignmentType.CENTER,
// // // //                   width: 1500,
// // // //                 }),
// // // //                 createCell("Amount", {
// // // //                   bold: true,
// // // //                   fontSize: 14,
// // // //                   alignment: AlignmentType.CENTER,
// // // //                   width: 1500,
// // // //                 }),
// // // //               ]),
// // // //         ],
// // // //       }),
// // // //       // Service rows
// // // //       ...serviceRows,
// // // //       // Empty row for spacing
// // // //       new TableRow({
// // // //         height: { value: 200, rule: HeightRule.ATLEAST },
// // // //         children: [
// // // //           createEmptyCell(6000),
// // // //           createEmptyCell(1500),
// // // //           createEmptyCell(1500),
// // // //           createEmptyCell(1500),
// // // //           createEmptyCell(1500),
// // // //         ],
// // // //       }),
// // // //       // Tax rows
// // // //       ...taxRows,
// // // //       // Total row
// // // //       new TableRow({
// // // //         children: [
// // // //           createCell(`Rupees: ${convertNumberToIndianWords(totalAmount)}`, {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             width: 7500,
// // // //             colSpan: 4,
// // // //           }),
// // // //           createCell("TOTAL", {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.CENTER,
// // // //             width: 1500,
// // // //           }),
// // // //           createCell(formatNumber(totalAmount), {
// // // //             bold: true,
// // // //             fontSize: 14,
// // // //             alignment: AlignmentType.RIGHT,
// // // //             width: 1500,
// // // //           }),
// // // //         ],
// // // //       }),
// // // //     ],
// // // //   });

// // // //   // Helper function to convert number to Indian words
// // // //   function convertNumberToIndianWords(amount: number): string {
// // // //     // Simplified version - you might want to use a more comprehensive library
// // // //     let rupees = Math.floor(amount);
// // // //     const paise = Math.round((amount - rupees) * 100);

// // // //     if (rupees === 0) return "Zero";

// // // //     const ones = [
// // // //       "",
// // // //       "One",
// // // //       "Two",
// // // //       "Three",
// // // //       "Four",
// // // //       "Five",
// // // //       "Six",
// // // //       "Seven",
// // // //       "Eight",
// // // //       "Nine",
// // // //     ];
// // // //     const teens = [
// // // //       "Ten",
// // // //       "Eleven",
// // // //       "Twelve",
// // // //       "Thirteen",
// // // //       "Fourteen",
// // // //       "Fifteen",
// // // //       "Sixteen",
// // // //       "Seventeen",
// // // //       "Eighteen",
// // // //       "Nineteen",
// // // //     ];
// // // //     const tens = [
// // // //       "",
// // // //       "",
// // // //       "Twenty",
// // // //       "Thirty",
// // // //       "Forty",
// // // //       "Fifty",
// // // //       "Sixty",
// // // //       "Seventy",
// // // //       "Eighty",
// // // //       "Ninety",
// // // //     ];

// // // //     let words = "";

// // // //     if (rupees >= 10000000) {
// // // //       words +=
// // // //         convertNumberToIndianWords(Math.floor(rupees / 10000000)) + " Crore ";
// // // //       rupees %= 10000000;
// // // //     }

// // // //     if (rupees >= 100000) {
// // // //       words +=
// // // //         convertNumberToIndianWords(Math.floor(rupees / 100000)) + " Lakh ";
// // // //       rupees %= 100000;
// // // //     }

// // // //     if (rupees >= 1000) {
// // // //       words +=
// // // //         convertNumberToIndianWords(Math.floor(rupees / 1000)) + " Thousand ";
// // // //       rupees %= 1000;
// // // //     }

// // // //     if (rupees >= 100) {
// // // //       words += ones[Math.floor(rupees / 100)] + " Hundred ";
// // // //       rupees %= 100;
// // // //     }

// // // //     if (rupees >= 20) {
// // // //       words += tens[Math.floor(rupees / 10)] + " ";
// // // //       rupees %= 10;
// // // //     }

// // // //     if (rupees >= 10) {
// // // //       words += teens[rupees - 10] + " ";
// // // //       rupees = 0;
// // // //     }

// // // //     if (rupees > 0) {
// // // //       words += ones[rupees] + " ";
// // // //     }

// // // //     return words.trim();
// // // //   }

// // // //   // Create document sections - Test without tables first
// // // //   const children = [
// // // //     // Simple test first
// // // //     new Paragraph({
// // // //       children: [
// // // //         new TextRun({
// // // //           text: `Test Document - ${data.customer.customerName}`,
// // // //           bold: true,
// // // //           font: "Arial",
// // // //           size: 18,
// // // //         }),
// // // //       ],
// // // //       spacing: { after: 200 },
// // // //     }),

// // // //     // Multiple paragraphs test
// // // //     new Paragraph({
// // // //       children: [
// // // //         new TextRun({
// // // //           text: `Customer: ${data.customer.customerName}`,
// // // //           font: "Arial",
// // // //           size: 14,
// // // //         }),
// // // //       ],
// // // //     }),

// // // //     new Paragraph({
// // // //       children: [
// // // //         new TextRun({
// // // //           text: `SAP Code: ${data.sapCode}`,
// // // //           font: "Arial",
// // // //           size: 14,
// // // //         }),
// // // //       ],
// // // //     }),

// // // //     new Paragraph({
// // // //       children: [
// // // //         new TextRun({
// // // //           text: `Total Value: ₹${data.totalValue.toLocaleString("en-IN")}`,
// // // //           font: "Arial",
// // // //           size: 14,
// // // //         }),
// // // //       ],
// // // //     }),
// // // //   ];

// // // //   const doc = new Document({
// // // //     sections: [
// // // //       {
// // // //         properties: {},
// // // //         children,
// // // //       },
// // // //     ],
// // // //   });

// // // //   try {
// // // //     console.log("📝 Creating document structure...");
// // // //     console.log("🔧 Document created, children:", children.length);

// // // //     // Try to identify potential issues
// // // //     console.log("🔍 Checking table structure...");
// // // //     const tables = children.filter((child) => child instanceof Table);
// // // //     console.log("📊 Found tables:", tables.length);

// // // //     console.log("📝 Attempting to create buffer...");
// // // //     const buffer = await Packer.toBuffer(doc);
// // // //     console.log("✅ Document created successfully, size:", buffer.length);
// // // //     return new Blob([new Uint8Array(buffer)], {
// // // //       type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
// // // //     });
// // // //   } catch (error) {
// // // //     console.error("❌ Error creating DOCX document:", error);
// // // //     console.error("❌ Error stack:", error.stack);
// // // //     console.error("❌ Error details:", {
// // // //       message: error.message,
// // // //       name: error.name,
// // // //       stack: error.stack,
// // // //     });

// // // //     // Try to create a minimal document to test if the library works at all
// // // //     console.log("🧪 Testing with minimal document...");
// // // //     try {
// // // //       const minimalDoc = new Document({
// // // //         sections: [
// // // //           {
// // // //             properties: {},
// // // //             children: [
// // // //               new Paragraph({
// // // //                 children: [new TextRun({ text: "Test document" })],
// // // //               }),
// // // //             ],
// // // //           },
// // // //         ],
// // // //       });
// // // //       const minimalBuffer = await Packer.toBuffer(minimalDoc);
// // // //       console.log("✅ Minimal document works, size:", minimalBuffer.length);
// // // //     } catch (minimalError) {
// // // //       console.error("❌ Even minimal document fails:", minimalError);
// // // //     }

// // // //     throw error;
// // // //   }
// // // // };



// // // import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } from "docx";
// // // import { saveAs } from "file-saver";
// // // import { convertNumberToIndianWords } from "@/lib/utils";
// // // import { extractStateFromAddress } from "@/lib/utils";
// // // import { MergedInvoiceData } from "@/types/invoice";

// // // export const generateInvoiceDOCX = async (
// // //   data: MergedInvoiceData,
// // //   invoiceType: "godown" | "main" | "freight",
// // //   formatNumber: (amount: number) => string
// // // ): Promise<Blob> => {
// // //   const customerState = extractStateFromAddress(data.customer.address);
// // //   const isInterState = ["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState);

// // //   const doc = new Document({
// // //     sections: [
// // //       {
// // //         children: [
// // //           new Paragraph({
// // //             text: "TAX INVOICE",
// // //             alignment: AlignmentType.RIGHT,
// // //             bold: true,
// // //           }),

// // //           new Paragraph({
// // //             text: data.customer.customerName,
// // //             bold: true,
// // //             spacing: { after: 150 },
// // //           }),

// // //           new Paragraph({
// // //             text: `Address: ${data.customer.address}`,
// // //           }),

// // //           new Paragraph({
// // //             text: `GSTIN: ${data.customer.gstin || "N/A"} | PAN: ${data.customer.pan || "N/A"}`,
// // //           }),

// // //           new Paragraph({
// // //             text: `Date: ${new Date().toLocaleDateString("en-IN")}`,
// // //             spacing: { after: 300 },
// // //           }),

// // //           new Paragraph({
// // //             text: `Service Description: ${
// // //               invoiceType === "godown"
// // //                 ? "Rental Services of Warehouse / Godown"
// // //                 : invoiceType === "main"
// // //                 ? "Loading / Unloading / Local Transportation Services"
// // //                 : "Secondary Freight"
// // //             }`,
// // //             bold: true,
// // //             spacing: { before: 200, after: 200 },
// // //           }),

// // //           // Table Section
// // //           createInvoiceTable(data, invoiceType, isInterState, formatNumber),

// // //           new Paragraph({
// // //             text: `Amount in Words: ${convertNumberToIndianWords(
// // //               calculateTotal(data, invoiceType, isInterState)
// // //             )}`,
// // //             spacing: { before: 300 },
// // //           }),

// // //           new Paragraph({
// // //             text: "TAX PAYABLE UNDER REVERSE CHARGE : NO",
// // //             bold: true,
// // //             spacing: { before: 300 },
// // //           }),

// // //           new Paragraph({
// // //             text: "Signature / Digital Signature",
// // //             alignment: AlignmentType.RIGHT,
// // //             bold: true,
// // //             spacing: { before: 200 },
// // //           }),
// // //         ],
// // //       },
// // //     ],
// // //   });

// // //   const blob = await Packer.toBlob(doc);
// // //   return blob;
// // // };

// // // // Helper to calculate total
// // // function calculateTotal(data, type, isInterState) {
// // //   let subtotal = 0;
// // //   switch (type) {
// // //     case "godown":
// // //       subtotal = data.godownAmount || 0;
// // //       break;
// // //     case "main":
// // //       subtotal =
// // //         (data.loadingAmount || 0) +
// // //         (data.unloadingAmount || 0) +
// // //         (data.localAmount || 0);
// // //       break;
// // //     case "freight":
// // //       subtotal = data.secondaryFreightAmount || 0;
// // //       break;
// // //   }
// // //   const tax = subtotal * 0.18;
// // //   return subtotal + tax;
// // // }

// // // // ✅ FIXED VERSION
// // // function createInvoiceTable(data, type, isInterState, formatNumber) {
// // //   const rows = [];

// // //   // Header row
// // //   rows.push(
// // //     new TableRow({
// // //       children: [
// // //         new TableCell({
// // //           children: [new Paragraph({ text: "Description of Services", bold: true })],
// // //           width: { size: 40, type: WidthType.PERCENTAGE },
// // //         }),
// // //         new TableCell({
// // //           children: [new Paragraph({ text: "SAC Code", bold: true })],
// // //         }),
// // //         new TableCell({
// // //           children: [new Paragraph({ text: "Amount (₹)", bold: true })],
// // //         }),
// // //       ],
// // //     })
// // //   );

// // //   // Build service rows
// // //   const serviceRows = [];

// // //   if (type === "godown") {
// // //     serviceRows.push([
// // //       "Rental Services of Warehouse / Godown",
// // //       "997212",
// // //       data.godownAmount || 0,
// // //     ]);
// // //   } else if (type === "main") {
// // //     serviceRows.push(["Loading Charges", "996511", data.loadingAmount || 0]);
// // //     serviceRows.push(["Unloading Charges", "996512", data.unloadingAmount || 0]);
// // //     serviceRows.push(["Local Transportation", "996513", data.localAmount || 0]);
// // //   } else {
// // //     serviceRows.push([
// // //       "Secondary Freight",
// // //       "996511",
// // //       data.secondaryFreightAmount || 0,
// // //     ]);
// // //   }

// // //   // Add each service row
// // //   for (const [desc, sac, amount] of serviceRows) {
// // //     rows.push(
// // //       new TableRow({
// // //         children: [
// // //           new TableCell({
// // //             children: [new Paragraph({ text: desc })],
// // //           }),
// // //           new TableCell({
// // //             children: [new Paragraph({ text: sac })],
// // //           }),
// // //           new TableCell({
// // //             children: [
// // //               new Paragraph({
// // //                 text: formatNumber(amount),
// // //                 alignment: AlignmentType.RIGHT,
// // //               }),
// // //             ],
// // //           }),
// // //         ],
// // //       })
// // //     );
// // //   }

// // //   // Subtotal
// // //   const subtotal = serviceRows.reduce((sum, [, , amt]) => sum + amt, 0);

// // //   // Tax rows
// // //   if (isInterState) {
// // //     rows.push(
// // //       new TableRow({
// // //         children: [
// // //           new TableCell({
// // //             children: [new Paragraph({ text: "IGST @18%" })],
// // //           }),
// // //           new TableCell({ children: [new Paragraph({ text: "" })] }),
// // //           new TableCell({
// // //             children: [
// // //               new Paragraph({
// // //                 text: formatNumber(subtotal * 0.18),
// // //                 alignment: AlignmentType.RIGHT,
// // //               }),
// // //             ],
// // //           }),
// // //         ],
// // //       })
// // //     );
// // //   } else {
// // //     rows.push(
// // //       new TableRow({
// // //         children: [
// // //           new TableCell({
// // //             children: [new Paragraph({ text: "CGST @9%" })],
// // //           }),
// // //           new TableCell({ children: [new Paragraph({ text: "" })] }),
// // //           new TableCell({
// // //             children: [
// // //               new Paragraph({
// // //                 text: formatNumber(subtotal * 0.09),
// // //                 alignment: AlignmentType.RIGHT,
// // //               }),
// // //             ],
// // //           }),
// // //         ],
// // //       })
// // //     );
// // //     rows.push(
// // //       new TableRow({
// // //         children: [
// // //           new TableCell({
// // //             children: [new Paragraph({ text: "SGST @9%" })],
// // //           }),
// // //           new TableCell({ children: [new Paragraph({ text: "" })] }),
// // //           new TableCell({
// // //             children: [
// // //               new Paragraph({
// // //                 text: formatNumber(subtotal * 0.09),
// // //                 alignment: AlignmentType.RIGHT,
// // //               }),
// // //             ],
// // //           }),
// // //         ],
// // //       })
// // //     );
// // //   }

// // //   // Total
// // //   rows.push(
// // //     new TableRow({
// // //       children: [
// // //         new TableCell({
// // //           children: [new Paragraph({ text: "TOTAL", bold: true })],
// // //         }),
// // //         new TableCell({ children: [new Paragraph({ text: "" })] }),
// // //         new TableCell({
// // //           children: [
// // //             new Paragraph({
// // //               text: formatNumber(subtotal + subtotal * 0.18),
// // //               bold: true,
// // //               alignment: AlignmentType.RIGHT,
// // //             }),
// // //           ],
// // //         }),
// // //       ],
// // //     })
// // //   );

// // //   return new Table({
// // //     rows,
// // //     width: { size: 100, type: WidthType.PERCENTAGE },
// // //   });
// // // }

// // import {
// //   Document,
// //   Packer,
// //   Paragraph,
// //   TextRun,
// //   Table,
// //   TableRow,
// //   TableCell,
// //   AlignmentType,
// //   WidthType,
// //   BorderStyle,
// // } from "docx";
// // import { convertNumberToIndianWords, extractStateFromAddress } from "@/lib/utils";
// // import type { MergedInvoiceData } from "@/types/invoice";

// // // Helper to create consistent paragraphs
// // const createParagraph = (
// //   text: string,
// //   options: {
// //     bold?: boolean;
// //     alignment?: AlignmentType;
// //     spacing?: { before?: number; after?: number };
// //     size?: number; // half-points (20 = 10pt)
// //     color?: string; // hex
// //   } = {}
// // ): Paragraph => {
// //   const { bold = false, alignment = AlignmentType.LEFT, spacing, size = 20, color } = options;
// //   return new Paragraph({
// //     alignment,
// //     spacing,
// //     children: [
// //       new TextRun({
// //         text,
// //         bold,
// //         font: "Arial",
// //         size,
// //         color: color || "000000",
// //       }),
// //     ],
// //   });
// // };

// // // Fixed Jubilant addresses by state (from your PDF logic)
// // const getJubilantDetails = (customerState: string) => {
// //   const base = {
// //     name: "Jubilant Agri & Consumer Products Ltd.",
// //     gstin: "06AADCC4657M1Z7",
// //   };

// //   if (["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState)) {
// //     return {
// //       ...base,
// //       address: [
// //         "ADD:- Ground Floor, 1233-1235,1243, Kapasan road,",
// //         "Village Singhpur, Tehsil Kapasan, Chittorgarh, Rajasthan, 312207",
// //       ],
// //       gstin: "08AADCC4657M1Z9",
// //     };
// //   }

// //   return {
// //     ...base,
// //     address: [
// //       "ADD:- NH-24, JUBILANT AGRI AND CONSUMER PRODUCTS LIMITED UNIT-I,",
// //       "BHARTIAGRAM, GAJRAULA, Amroha, Uttar Pradesh, 244223",
// //     ],
// //   };
// // };

// // export const generateInvoiceDOCX = async (
// //   data: MergedInvoiceData,
// //   invoiceType: "godown" | "main" | "freight",
// //   formatNumber: (amount: number) => string
// // ): Promise<Blob> => {
// //   const customerState = extractStateFromAddress(data.customer.address);
// //   const isInterState = ["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState);
// //   const jubilant = getJubilantDetails(customerState);

// //   // === 1. Header: "TAX" (red) + "INVOICE" (black) ===
// //   const headerParagraph = new Paragraph({
// //     children: [
// //       new TextRun({
// //         text: data.customer.customerName,
// //         bold: true,
// //         size: 18, // 9pt
// //         font: "Arial",
// //       }),
// //       new TextRun({
// //         text: "\t\t\t\t\t\t\t\t\t", // tab to push right
// //       }),
// //       new TextRun({
// //         text: "TAX",
// //         bold: true,
// //         size: 22, // 11pt
// //         font: "Arial",
// //         color: "FF0000",
// //       }),
// //       new TextRun({
// //         text: " INVOICE",
// //         bold: true,
// //         size: 22,
// //         font: "Arial",
// //         color: "000000",
// //       }),
// //     ],
// //     spacing: { after: 100 },
// //   });

// //   // === 2. Customer Details (Left-aligned, labeled) ===
// //   const customerDetails = [
// //     createParagraph(`Add: ${data.customer.address}`, { size: 14 }), // 7pt
// //     createParagraph(
// //       `GSTIN: ${data.customer.gstin || "N/A"}   PAN: ${data.customer.pan || "N/A"}`,
// //       { size: 14 }
// //     ),
// //     createParagraph(
// //       `EMAIL: ${data.customer.email || "N/A"}   MOB: ${data.customer.mobile || "N/A"}`,
// //       { size: 14 }
// //     ),
// //   ];

// //   // === 3. Jubilant (Sender) Details ===
// //   const jubilantDetails = [
// //     createParagraph("To:", { bold: true, size: 14 }),
// //     createParagraph(jubilant.name, { bold: true, size: 14 }),
// //     ...jubilant.address.map((line) => createParagraph(line, { size: 14 })),
// //     createParagraph(`GSTIN : ${jubilant.gstin}`, { bold: true, size: 14 }),
// //   ];

// //   // === 4. Service Description (Top-right) ===
// //   let serviceDesc = "";
// //   if (invoiceType === "godown") {
// //     serviceDesc = `Rental or Leasing services involving own or leased non - residential property for ${customerState}`;
// //   } else if (invoiceType === "main") {
// //     serviceDesc = "Clearing & Forwording Charges";
// //   } // freight: blank

// //   const serviceDescParagraph = createParagraph(serviceDesc, {
// //     alignment: AlignmentType.RIGHT,
// //     size: 14,
// //     spacing: { before: 100, after: 200 },
// //   });

// //   // === 5. Table Data ===
// //   let tableHeaders: string[] = [];
// //   let serviceRows: any[] = [];
// //   let totalAmountBeforeTax = 0;

// //   if (invoiceType === "freight") {
// //     tableHeaders = ["Service Description", "HSN / SAC", "Amount"];
// //     serviceRows.push([
// //       "Secondary Freight",
// //       "996511",
// //       formatNumber(data.freightBalance),
// //     ]);
// //     totalAmountBeforeTax = data.freightBalance;
// //   } else if (invoiceType === "main") {
// //     tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
// //     serviceRows.push(
// //       [
// //         "Loading Charges",
// //         "996519",
// //         formatNumber(data.loadingCharges / 75),
// //         "75.00",
// //         formatNumber(data.loadingCharges),
// //       ],
// //       [
// //         "Unloading Charges",
// //         "996519",
// //         formatNumber(data.unloadingCharges / 75),
// //         "75.00",
// //         formatNumber(data.unloadingCharges),
// //       ],
// //       [
// //         "Local Transportation",
// //         "996713",
// //         formatNumber(data.localTransportation / 200),
// //         "200.00",
// //         formatNumber(data.localTransportation),
// //       ]
// //     );
// //     totalAmountBeforeTax =
// //       data.loadingCharges + data.unloadingCharges + data.localTransportation;
// //   } else {
// //     tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
// //     serviceRows.push([
// //       "Rental or Leasing services involving own or leased non - residential property",
// //       "997212",
// //       formatNumber(data.godownRent / 100),
// //       "100.00",
// //       formatNumber(data.godownRent),
// //     ]);
// //     totalAmountBeforeTax = data.godownRent;
// //   }

// //   // Tax calculation
// //   let cgst = 0, sgst = 0, igst = 0, totalAmount = 0;
// //   if (isInterState) {
// //     igst = totalAmountBeforeTax * 0.18;
// //     totalAmount = totalAmountBeforeTax + igst;
// //   } else {
// //     cgst = sgst = totalAmountBeforeTax * 0.09;
// //     totalAmount = totalAmountBeforeTax + cgst + sgst;
// //   }

// //   // Build table rows
// //   const tableRows: TableRow[] = [];

// //   // Header
// //   tableRows.push(
// //     new TableRow({
// //       children: tableHeaders.map((header) =>
// //         new TableCell({
// //           children: [createParagraph(header, { bold: true, size: 14 })],
// //           shading: { fill: "FFFFFF" },
// //         })
// //       ),
// //     })
// //   );

// //   // Service rows
// //   serviceRows.forEach((row) => {
// //     tableRows.push(
// //       new TableRow({
// //         children: row.map((cell, idx) =>
// //           new TableCell({
// //             children: [
// //               createParagraph(cell, {
// //                 alignment:
// //                   idx === row.length - 1 ? AlignmentType.RIGHT : AlignmentType.LEFT,
// //                 size: 14,
// //               }),
// //             ],
// //           })
// //         ),
// //       })
// //     );
// //   });

// //   // Empty row (spacing)
// //   tableRows.push(
// //     new TableRow({
// //       children: Array(tableHeaders.length).fill(0).map(() => new TableCell({ children: [new Paragraph("")], height: 80 })),
// //     })
// //   );

// //   // Tax rows
// //   if (isInterState) {
// //     tableRows.push(
// //       new TableRow({
// //         children: [
// //           new TableCell({ children: [createParagraph("IGST @ 18%", { size: 14 })] }),
// //           ...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")]})),
// //           new TableCell({ children: [createParagraph(formatNumber(igst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
// //         ],
// //       })
// //     );
// //   } else {
// //     tableRows.push(
// //       new TableRow({
// //         children: [
// //           new TableCell({ children: [createParagraph("CGST @ 9%", { size: 14 })] }),
// //           ...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")]})),
// //           new TableCell({ children: [createParagraph(formatNumber(cgst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
// //         ],
// //       })
// //     );
// //     tableRows.push(
// //       new TableRow({
// //         children: [
// //           new TableCell({ children: [createParagraph("SGST @ 9%", { size: 14 })] }),
// //           ...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")]})),
// //           new TableCell({ children: [createParagraph(formatNumber(sgst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
// //         ],
// //       })
// //     );
// //   }

// //   // Total row with amount in words
// //   const amountInWords = `Rupees: ${convertNumberToIndianWords(totalAmount)}`;
// //   tableRows.push(
// //     new TableRow({
// //       children: [
// //         new TableCell({
// //           children: [createParagraph(amountInWords, { size: 14 })],
// //           columnSpan: tableHeaders.length - 2,
// //         }),
// //         new TableCell({ children: [createParagraph("TOTAL", { bold: true, size: 14 })] }),
// //         new TableCell({
// //           children: [createParagraph(formatNumber(totalAmount), { bold: true, alignment: AlignmentType.RIGHT, size: 14 })],
// //         }),
// //       ],
// //     })
// //   );

// //   const table = new Table({
// //     rows: tableRows,
// //     width: { size: 100, type: WidthType.PERCENTAGE },
// //     borders: {
// //       top: { style: BorderStyle.SINGLE, size: 1 },
// //       bottom: { style: BorderStyle.SINGLE, size: 1 },
// //       left: { style: BorderStyle.SINGLE, size: 1 },
// //       right: { style: BorderStyle.SINGLE, size: 1 },
// //       insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
// //       insideVertical: { style: BorderStyle.SINGLE, size: 1 },
// //     },
// //   });

// //   // === 6. Footer ===
// //   const footer = [
// //     createParagraph("TAX PAYABLE UNDER REVERSE CHARGE : NO", {
// //       bold: true,
// //       color: "FF0000",
// //       size: 14,
// //       spacing: { before: 200 },
// //     }),
// //     createParagraph("SIGNATURE/ DIGITAL SIGNATURE", {
// //       bold: true,
// //       alignment: AlignmentType.RIGHT,
// //       color: "FF0000",
// //       size: 14,
// //       spacing: { before: 100 },
// //     }),
// //   ];

// //   // === Assemble Document ===
// //   const doc = new Document({
// //     sections: [
// //       {
// //         children: [
// //           headerParagraph,
// //           ...customerDetails,
// //           ...jubilantDetails,
// //           serviceDescParagraph,
// //           table,
// //           ...footer,
// //         ],
// //       },
// //     ],
// //   });

// //   return await Packer.toBlob(doc);
// // };


// import {
//   Document,
//   Packer,
//   Paragraph,
//   TextRun,
//   Table,
//   TableRow,
//   TableCell,
//   AlignmentType,
//   WidthType,
//   BorderStyle,
// } from "docx";
// import { convertNumberToIndianWords, extractStateFromAddress } from "@/lib/utils";
// import type { MergedInvoiceData } from "@/types/invoice";

// // Helper to create styled paragraphs
// const createParagraph = (
//   text: string,
//   options: {
//     bold?: boolean;
//     alignment?: AlignmentType;
//     size?: number;
//     color?: string;
//     spacing?: { before?: number; after?: number };
//   } = {}
// ): Paragraph => {
//   const {
//     bold = false,
//     alignment = AlignmentType.LEFT,
//     size = 18, // default ~9pt
//     color = "000000",
//     spacing = { before: 0, after: 0 },
//   } = options;

//   return new Paragraph({
//     alignment,
//     spacing,
//     children: [
//       new TextRun({
//         text,
//         bold,
//         size,
//         color,
//         font: "Arial",
//       }),
//     ],
//   });
// };

// // Jubilant address logic (like your PDF)
// const getJubilantDetails = (customerState: string) => {
//   const base = {
//     name: "Jubilant Agri & Consumer Products Ltd.",
//     gstin: "06AADCC4657M1Z7",
//   };

//   if (["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState)) {
//     return {
//       ...base,
//       address: [
//         "ADD:- Ground Floor, 1233-1235,1243, Kapasan road,",
//         "Village Singhpur, Tehsil Kapasan, Chittorgarh, Rajasthan, 312207",
//       ],
//       gstin: "08AADCC4657M1Z9",
//     };
//   }

//   return {
//     ...base,
//     address: [
//       "ADD:- NH-24, JUBILANT AGRI AND CONSUMER PRODUCTS LIMITED UNIT-I,",
//       "BHARTIAGRAM, GAJRAULA, Amroha, Uttar Pradesh, 244223",
//     ],
//   };
// };

// export const generateInvoiceDOCX = async (
//   data: MergedInvoiceData,
//   invoiceType: "godown" | "main" | "freight",
//   formatNumber: (amount: number) => string
// ): Promise<Blob> => {
//   const customerState = extractStateFromAddress(data.customer.address);
//   const isInterState = ["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState);
//   const jubilant = getJubilantDetails(customerState);

//   // Leave invoice number and date blank
//   const invoiceNo = "";
//   const invoiceDate = "";

//   // === HEADER ===
//   const leftHeaderCellChildren: Paragraph[] = [
//     createParagraph(data.customer.customerName, { bold: true, size: 18 }),
//   ];

//   const rightHeaderChildren: Paragraph[] = [
//     new Paragraph({
//       children: [
//         new TextRun({
//           text: "TAX",
//           bold: true,
//           color: "FF0000",
//           size: 22,
//           font: "Arial",
//         }),
//         new TextRun({
//           text: " INVOICE",
//           bold: true,
//           color: "000000",
//           size: 22,
//           font: "Arial",
//         }),
//       ],
//       alignment: AlignmentType.RIGHT,
//       spacing: { after: 120 },
//     }),
//     createParagraph(`INVOICE NO : ${invoiceNo}`, { alignment: AlignmentType.RIGHT, size: 14 }),
//     createParagraph(`DATE : ${invoiceDate}`, { alignment: AlignmentType.RIGHT, size: 14, spacing: { after: 100 } }),
//   ];

//   // Header table with no top border or inside lines
//   const headerTable = new Table({
//     rows: [
//       new TableRow({
//         children: [
//           new TableCell({
//             children: leftHeaderCellChildren,
//             width: { size: 65, type: WidthType.PERCENTAGE },
//             margins: { top: 0, bottom: 0, left: 0, right: 0 },
//             borders: {
//               top: { style: BorderStyle.NONE },
//               bottom: { style: BorderStyle.NONE },
//               left: { style: BorderStyle.NONE },
//               right: { style: BorderStyle.NONE },
//             },
//           }),
//           new TableCell({
//             children: rightHeaderChildren,
//             width: { size: 35, type: WidthType.PERCENTAGE },
//             margins: { top: 0, bottom: 0, left: 0, right: 0 },
//             borders: {
//               top: { style: BorderStyle.NONE },
//               bottom: { style: BorderStyle.NONE },
//               left: { style: BorderStyle.NONE },
//               right: { style: BorderStyle.NONE },
//             },
//           }),
//         ],
//       }),
//     ],
//     width: { size: 100, type: WidthType.PERCENTAGE },
//     borders: {
//       top: { style: BorderStyle.NONE },
//       bottom: { style: BorderStyle.NONE },
//       left: { style: BorderStyle.NONE },
//       right: { style: BorderStyle.NONE },
//       insideHorizontal: { style: BorderStyle.NONE },
//       insideVertical: { style: BorderStyle.NONE },
//     },
//   });

//   // === CUSTOMER DETAILS ===
//   const addressPara = createParagraph(`Add: ${data.customer.address}`, { size: 14 });
//   const gstPan = createParagraph(
//     `GSTIN: ${data.customer.gstin || "N/A"}     PAN: ${data.customer.pan || "N/A"}`,
//     { size: 14 }
//   );
//   const emailMob = createParagraph(
//     `EMAIL: ${data.customer.email || "N/A"}     MOB: ${data.customer.mobile || "N/A"}`,
//     { size: 14, spacing: { after: 100 } }
//   );

//   // === JUBILANT DETAILS ===
//   const jubilantParas = [
//     createParagraph("To:", { bold: true, size: 14 }),
//     createParagraph(jubilant.name, { bold: true, size: 14 }),
//     ...jubilant.address.map((line) => createParagraph(line, { size: 14 })),
//     createParagraph(`GSTIN: ${jubilant.gstin}`, { bold: true, size: 14, spacing: { after: 200 } }),
//   ];

//   // === SERVICE DESCRIPTION ===
//   let serviceDesc = "";
//   if (invoiceType === "godown")
//     serviceDesc = `Rental or Leasing services involving own or leased non - residential property for ${customerState}`;
//   else if (invoiceType === "main")
//     serviceDesc = "Clearing & Forwording Charges";
//   else serviceDesc = "";

//   const serviceDescPara = createParagraph(serviceDesc, {
//     alignment: AlignmentType.RIGHT,
//     size: 14,
//     spacing: { before: 100, after: 200 },
//   });

//   // === TABLE DATA ===
//   let tableHeaders: string[] = [];
//   let serviceRows: any[] = [];
//   let totalAmountBeforeTax = 0;

//   if (invoiceType === "freight") {
//     tableHeaders = ["Service Description", "HSN / SAC", "Amount"];
//     serviceRows.push(["Secondary Freight", "996511", formatNumber(data.freightBalance)]);
//     totalAmountBeforeTax = data.freightBalance;
//   } else if (invoiceType === "main") {
//     tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
//     serviceRows.push(
//       ["Loading Charges", "996519", formatNumber(data.loadingCharges / 75), "75.00", formatNumber(data.loadingCharges)],
//       ["Unloading Charges", "996519", formatNumber(data.unloadingCharges / 75), "75.00", formatNumber(data.unloadingCharges)],
//       ["Local Transportation", "996713", formatNumber(data.localTransportation / 200), "200.00", formatNumber(data.localTransportation)]
//     );
//     totalAmountBeforeTax =
//       data.loadingCharges + data.unloadingCharges + data.localTransportation;
//   } else {
//     tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
//     serviceRows.push([
//       "Rental or Leasing services involving own or leased non - residential property",
//       "997212",
//       formatNumber(data.godownRent / 100),
//       "100.00",
//       formatNumber(data.godownRent),
//     ]);
//     totalAmountBeforeTax = data.godownRent;
//   }

//   // === TAX LOGIC ===
//   let cgst = 0,
//     sgst = 0,
//     igst = 0,
//     totalAmount = 0;
//   if (isInterState) {
//     igst = totalAmountBeforeTax * 0.18;
//     totalAmount = totalAmountBeforeTax + igst;
//   } else {
//     cgst = sgst = totalAmountBeforeTax * 0.09;
//     totalAmount = totalAmountBeforeTax + cgst + sgst;
//   }

//   // === TABLE CONSTRUCTION ===
//   const tableRows: TableRow[] = [];

//   // Header
//   tableRows.push(
//     new TableRow({
//       children: tableHeaders.map(
//         (h) =>
//           new TableCell({
//             children: [createParagraph(h, { bold: true, size: 14, alignment: AlignmentType.CENTER })],
//           })
//       ),
//     })
//   );

//   // Service Rows
//   serviceRows.forEach((row) => {
//     tableRows.push(
//       new TableRow({
//         children: row.map((cell: string, idx: number) =>
//           new TableCell({
//             children: [
//               createParagraph(cell, {
//                 alignment:
//                   idx === row.length - 1
//                     ? AlignmentType.RIGHT
//                     : idx === 0
//                     ? AlignmentType.LEFT
//                     : AlignmentType.CENTER,
//                 size: 14,
//               }),
//             ],
//           })
//         ),
//       })
//     );
//   });

//   // Empty spacing row
//   tableRows.push(
//     new TableRow({
//       children: tableHeaders.map(() => new TableCell({ children: [new Paragraph("")] })),
//     })
//   );

//   // Tax rows
//   if (isInterState) {
//     tableRows.push(
//       new TableRow({
//         children: [
//           new TableCell({ children: [createParagraph("IGST @ 18%", { size: 14 })] }),
//           ...Array(tableHeaders.length - 2)
//             .fill(0)
//             .map(() => new TableCell({ children: [new Paragraph("")] })),
//           new TableCell({
//             children: [createParagraph(formatNumber(igst), { alignment: AlignmentType.RIGHT, size: 14 })],
//           }),
//         ],
//       })
//     );
//   } else {
//     tableRows.push(
//       new TableRow({
//         children: [
//           new TableCell({ children: [createParagraph("CGST @ 9%", { size: 14 })] }),
//           ...Array(tableHeaders.length - 2)
//             .fill(0)
//             .map(() => new TableCell({ children: [new Paragraph("")] })),
//           new TableCell({
//             children: [createParagraph(formatNumber(cgst), { alignment: AlignmentType.RIGHT, size: 14 })],
//           }),
//         ],
//       }),
//       new TableRow({
//         children: [
//           new TableCell({ children: [createParagraph("SGST @ 9%", { size: 14 })] }),
//           ...Array(tableHeaders.length - 2)
//             .fill(0)
//             .map(() => new TableCell({ children: [new Paragraph("")] })),
//           new TableCell({
//             children: [createParagraph(formatNumber(sgst), { alignment: AlignmentType.RIGHT, size: 14 })],
//           }),
//         ],
//       })
//     );
//   }

//   // Total Row
//   const amountInWords = `Rupees: ${convertNumberToIndianWords(totalAmount)}`;
//   tableRows.push(
//     new TableRow({
//       children: [
//         new TableCell({
//           children: [createParagraph(amountInWords, { size: 14 })],
//           columnSpan: tableHeaders.length - 2,
//         }),
//         new TableCell({ children: [createParagraph("TOTAL", { bold: true, size: 14 })] }),
//         new TableCell({
//           children: [
//             createParagraph(formatNumber(totalAmount), {
//               bold: true,
//               alignment: AlignmentType.RIGHT,
//               size: 14,
//             }),
//           ],
//         }),
//       ],
//     })
//   );

//   const table = new Table({
//     rows: tableRows,
//     width: { size: 100, type: WidthType.PERCENTAGE },
//     borders: {
//       top: { style: BorderStyle.SINGLE, size: 1 },
//       bottom: { style: BorderStyle.SINGLE, size: 1 },
//       left: { style: BorderStyle.SINGLE, size: 1 },
//       right: { style: BorderStyle.SINGLE, size: 1 },
//       insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
//       insideVertical: { style: BorderStyle.SINGLE, size: 1 },
//     },
//   });

//   // === FOOTER ===
//   const footer = [
//     createParagraph("TAX PAYABLE UNDER REVERSE CHARGE : NO", {
//       bold: true,
//       color: "FF0000",
//       size: 14,
//       spacing: { before: 200 },
//     }),
//     createParagraph("SIGNATURE/ DIGITAL SIGNATURE", {
//       bold: true,
//       alignment: AlignmentType.RIGHT,
//       color: "FF0000",
//       size: 14,
//       spacing: { before: 100 },
//     }),
//   ];

//   // === COMPILE DOCUMENT ===
//   const doc = new Document({
//     sections: [
//       {
//         children: [
//           headerTable,
//           addressPara,
//           gstPan,
//           emailMob,
//           ...jubilantParas,
//           serviceDescPara,
//           table,
//           ...footer,
//         ],
//       },
//     ],
//   });

//   return await Packer.toBlob(doc);
// };



// new code to fix formatting issues
import {
Document,
Packer,
Paragraph,
TextRun,
Table,
TableRow,
TableCell,
AlignmentType,
WidthType,
BorderStyle,
} from "docx";
import { convertNumberToIndianWords, extractStateFromAddress } from "@/lib/utils";
import type { MergedInvoiceData } from "@/types/invoice";

// Helper to create styled paragraphs
const createParagraph = (
text: string,
options: {
bold?: boolean;
alignment?: AlignmentType;
size?: number;
color?: string;
spacing?: { before?: number; after?: number };
} = {}
): Paragraph => {
const {
bold = false,
alignment = AlignmentType.LEFT,
size = 18, // default ~9pt
color = "000000",
spacing = { before: 0, after: 0 },
} = options;

return new Paragraph({
alignment,
spacing,
children: [
new TextRun({
text,
bold,
size,
color,
font: "Arial",
}),
],
});
};

// Jubilant address logic (like your PDF)
const getJubilantDetails = (customerState: string) => {
const base = {
name: "Jubilant Agri & Consumer Products Ltd.",
gstin: "06AADCC4657M1Z7",
};

if (["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState)) {
return {
...base,
address: [
"ADD:- Ground Floor, 1233-1235,1243, Kapasan road,",
"Village Singhpur, Tehsil Kapasan, Chittorgarh, Rajasthan, 312207",
],
gstin: "08AADCC4657M1Z9",
};
}

return {
...base,
address: [
"ADD:- NH-24, JUBILANT AGRI AND CONSUMER PRODUCTS LIMITED UNIT-I,",
"BHARTIAGRAM, GAJRAULA, Amroha, Uttar Pradesh, 244223",
],
};
};

export const generateInvoiceDOCX = async (
data: MergedInvoiceData,
invoiceType: "godown" | "main" | "freight",
formatNumber: (amount: number) => string
): Promise<Blob> => {
const customerState = extractStateFromAddress(data.customer.address);
const isInterState = ["Maharashtra", "Gujarat", "Chhattisgarh", "Uttarakhand"].includes(customerState);
const jubilant = getJubilantDetails(customerState);

const invoiceNo = "";
const invoiceDate = "";

// === HEADER ===
const leftHeaderCellChildren: Paragraph[] = [
createParagraph(data.customer.customerName, { bold: true, size: 18 }),
];

const rightHeaderChildren: Paragraph[] = [
new Paragraph({
children: [
new TextRun({
text: "TAX",
bold: true,
color: "FF0000",
size: 22,
font: "Arial",
}),
new TextRun({
text: " INVOICE",
bold: true,
color: "000000",
size: 22,
font: "Arial",
}),
],
alignment: AlignmentType.RIGHT,
spacing: { after: 120 },
}),
createParagraph(`INVOICE NO : ${invoiceNo}`, { alignment: AlignmentType.RIGHT, size: 14 }),
createParagraph(`DATE : ${invoiceDate}`, { alignment: AlignmentType.RIGHT, size: 14, spacing: { after: 100 } }),
];

const headerTable = new Table({
rows: [
new TableRow({
children: [
new TableCell({
children: leftHeaderCellChildren,
width: { size: 65, type: WidthType.PERCENTAGE },
borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
}),
new TableCell({
children: rightHeaderChildren,
width: { size: 35, type: WidthType.PERCENTAGE },
borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
}),
],
}),
],
width: { size: 100, type: WidthType.PERCENTAGE },
borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
});

// === CUSTOMER DETAILS ===
const addressPara = createParagraph(`Add: ${data.customer.address}`, { size: 14, spacing: { before: 200, after: 100 } });
const gstPan = createParagraph(`GSTIN: ${data.customer.gstin || "N/A"}     PAN: ${data.customer.pan || "N/A"}`, { size: 14 });
const emailMob = createParagraph(`EMAIL: ${data.customer.email || "N/A"}     MOB: ${data.customer.mobile || "N/A"}`, { size: 14, spacing: { after: 300 } });

// === JUBILANT DETAILS ===
const jubilantParas = [
createParagraph("To:", { bold: true, size: 14 }),
createParagraph(jubilant.name, { bold: true, size: 14 }),
...jubilant.address.map((line) => createParagraph(line, { size: 14 })),
createParagraph(`GSTIN: ${jubilant.gstin}`, { bold: true, size: 14, spacing: { after: 300 } }),
];

// === SERVICE DESCRIPTION ===
let serviceDesc = "";
if (invoiceType === "godown")
serviceDesc = `Rental or Leasing services involving own or leased non - residential property for ${customerState}`;
else if (invoiceType === "main")
serviceDesc = "Clearing & Forwording Charges";
else serviceDesc = "";

const serviceDescPara = createParagraph(serviceDesc, {
alignment: AlignmentType.RIGHT,
size: 14,
spacing: { before: 300, after: 300 },
});

// === TABLE DATA ===
let tableHeaders: string[] = [];
let serviceRows: any[] = [];
let totalAmountBeforeTax = 0;

if (invoiceType === "freight") {
tableHeaders = ["Service Description", "HSN / SAC", "Amount"];
serviceRows.push(["Secondary Freight", "996511", formatNumber(data.freightBalance)]);
totalAmountBeforeTax = data.freightBalance;
} else if (invoiceType === "main") {
tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
serviceRows.push(
["Loading Charges", "996519", formatNumber(data.loadingCharges / 75), "75.00", formatNumber(data.loadingCharges)],
["Unloading Charges", "996519", formatNumber(data.unloadingCharges / 75), "75.00", formatNumber(data.unloadingCharges)],
["Local Transportation", "996713", formatNumber(data.localTransportation / 200), "200.00", formatNumber(data.localTransportation)]
);
totalAmountBeforeTax = data.loadingCharges + data.unloadingCharges + data.localTransportation;
} else {
tableHeaders = ["Service Description", "HSN / SAC", "Qty", "Rate", "Amount"];
serviceRows.push([
"Rental or Leasing services involving own or leased non - residential property",
"997212",
formatNumber(data.godownRent / 100),
"100.00",
formatNumber(data.godownRent),
]);
totalAmountBeforeTax = data.godownRent;
}

// === TAX LOGIC ===
let cgst = 0, sgst = 0, igst = 0, totalAmount = 0;
if (isInterState) {
igst = totalAmountBeforeTax * 0.18;
totalAmount = totalAmountBeforeTax + igst;
} else {
cgst = sgst = totalAmountBeforeTax * 0.09;
totalAmount = totalAmountBeforeTax + cgst + sgst;
}

// === TABLE CONSTRUCTION ===
const tableRows: TableRow[] = [];

// Header row
tableRows.push(
new TableRow({
children: tableHeaders.map((h) =>
new TableCell({
children: [createParagraph(h, { bold: true, size: 14, alignment: AlignmentType.CENTER })],
})
),
})
);

// Service rows
serviceRows.forEach((row) => {
tableRows.push(
new TableRow({
children: row.map((cell: string, idx: number) =>
new TableCell({
children: [
createParagraph(cell, {
alignment:
idx === row.length - 1
? AlignmentType.RIGHT
: idx === 0
? AlignmentType.LEFT
: AlignmentType.CENTER,
size: 14,
}),
],
})
),
})
);
});

// Empty row
tableRows.push(new TableRow({ children: tableHeaders.map(() => new TableCell({ children: [new Paragraph("")] })) }));

// Tax rows
if (isInterState) {
tableRows.push(
new TableRow({
children: [
new TableCell({ children: [createParagraph("IGST @ 18%", { size: 14 })] }),
...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")] })),
new TableCell({ children: [createParagraph(formatNumber(igst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
],
})
);
} else {
tableRows.push(
new TableRow({
children: [
new TableCell({ children: [createParagraph("CGST @ 9%", { size: 14 })] }),
...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")] })),
new TableCell({ children: [createParagraph(formatNumber(cgst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
],
}),
new TableRow({
children: [
new TableCell({ children: [createParagraph("SGST @ 9%", { size: 14 })] }),
...Array(tableHeaders.length - 2).fill(0).map(() => new TableCell({ children: [new Paragraph("")] })),
new TableCell({ children: [createParagraph(formatNumber(sgst), { alignment: AlignmentType.RIGHT, size: 14 })] }),
],
})
);
}

const amountInWords = `Rupees: ${convertNumberToIndianWords(totalAmount)}`;
tableRows.push(
new TableRow({
children: [
new TableCell({
children: [createParagraph(amountInWords, { size: 14 })],
columnSpan: tableHeaders.length - 2,
}),
new TableCell({ children: [createParagraph("TOTAL", { bold: true, size: 14 })] }),
new TableCell({
children: [createParagraph(formatNumber(totalAmount), { bold: true, alignment: AlignmentType.RIGHT, size: 14 })],
}),
],
})
);

const table = new Table({
rows: tableRows,
width: { size: 100, type: WidthType.PERCENTAGE },
borders: {
top: { style: BorderStyle.SINGLE, size: 1 },
bottom: { style: BorderStyle.SINGLE, size: 1 },
left: { style: BorderStyle.SINGLE, size: 1 },
right: { style: BorderStyle.SINGLE, size: 1 },
insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
insideVertical: { style: BorderStyle.SINGLE, size: 1 },
},
});

// === FOOTER ===
const blankSpace = new Paragraph({ text: "", spacing: { after: 400 } });
const footer = [
blankSpace,
blankSpace,
createParagraph("TAX PAYABLE UNDER REVERSE CHARGE : NO", {
bold: true,
color: "FF0000",
size: 14,
spacing: { before: 300 },
}),
createParagraph("SIGNATURE/ DIGITAL SIGNATURE", {
bold: true,
alignment: AlignmentType.RIGHT,
color: "FF0000",
size: 14,
spacing: { before: 200 },
}),
];

// === COMPILE DOCUMENT ===
const doc = new Document({
sections: [
{
properties: {
page: {
margin: {
top: 1440, // 1 inch
bottom: 1440,
left: 1080,
right: 1080,
},
},
},
children: [
headerTable,
addressPara,
gstPan,
emailMob,
...jubilantParas,
serviceDescPara,
table,
...footer,
],
},
],
});

return await Packer.toBlob(doc);
};
