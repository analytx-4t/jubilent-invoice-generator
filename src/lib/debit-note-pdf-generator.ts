import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import type { MergedInvoiceData } from "@/types/invoice";
import {
  convertNumberToIndianWords,
  extractStateFromAddress,
  otherMainStates,
} from "@/lib/utils";
import { JubilantDetails } from "@/components/JubilantDetails";

const generateReferenceDebitNotePDF = (
  data: MergedInvoiceData,
  invoiceType: "main" | "freight",
  formatNumber: (amount: number) => string,
  debitMonth: string
): Blob => {
  const doc = new jsPDF();
  const pageWidth = doc.internal.pageSize.getWidth();
  const isFreight = invoiceType === "freight";
  const subject = isFreight
    ? "Subject: Request for Issuance of Credit Note towards Secondary Freight support/Reimbursement"
    : "Subject: Request for Issuance of Credit Note towards Handling Charges support/Reimbursement";
  const requestText = isFreight
    ? "We request you to kindly issue a credit note towards the support/reimbursement on account of secondary freight expenses incurred by us for the onward movement of goods."
    : "We request you to kindly issue a credit note towards the support/reimbursement on account of Handling Charges incurred by us for the onward movement of goods.";
  const particulars = isFreight
    ? `Secondary Freight Support/Reimbursement for the month ${debitMonth}`
    : `Handling Charges support/reimbursement for the month ${debitMonth}`;
  const amount = isFreight
    ? data.freightBalance
    : (data.loadingCharges || 0) +
      (data.unloadingCharges || 0) +
      (data.localTransportation || 0);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(data.customer.customerName, pageWidth / 2, 16, { align: "center" });
  doc.text("ADDRESS: " + (data.customer.address || ""), pageWidth / 2, 28, {
    align: "center",
  });
  doc.text(
    "GSTIN: " +
      data.customer.gstin +
      "  PAN: " +
      data.customer.pan +
      "  MOB: " +
      data.customer.mobile,
    pageWidth / 2,
    38,
    { align: "center" }
  );

  doc.setFont("helvetica", "normal");
  doc.text("To,", 20, 58);
  doc.setFont("helvetica", "bold");
  doc.text("Dated:", pageWidth - 20, 58, { align: "right" });
  doc.setFont("helvetica", "normal");
  [
    "The Sales Head",
    "Jubilant Agri & Consumer Products Ltd.",
    "Plot No 142, Chimes, 3rd Floor, Sector 44,",
    "Gurugram Haryana-122003",
  ].forEach((line, index) => doc.text(line, 20, 67 + index * 5));

  doc.text(subject, 20, 105);
  doc.text("Dear sir/madam,", 20, 119);
  doc.text(doc.splitTextToSize(requestText, 175), 20, 129);

  doc.setFont("helvetica", "bold");
  doc.text("Particulars", 23, 151);
  doc.text("Amt. INR", pageWidth - 20, 151, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.text(particulars.split("\n"), 23, 160);
  doc.text(formatNumber(amount), pageWidth - 20, 160, { align: "right" });

  doc.setFont("helvetica", "bold");
  doc.text("Total", 23, 184);
  doc.text(formatNumber(amount), pageWidth - 20, 184, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.text("Thanking You", 20, 204);
  doc.text("Yours Faithfully", 20, 212);
  doc.text("Authorised Signatory", 20, 228);

  return doc.output("blob");
};

export const generateDebitNotePDF = (
  data: MergedInvoiceData,
  invoiceType: "godown" | "main" | "freight",
  formatNumber: (amount: number) => string,
  debitMonth :string
): Blob => {
  if ((invoiceType as string) === "main" || (invoiceType as string) === "freight") {
      return generateReferenceDebitNotePDF(
        data,
        invoiceType as "main" | "freight",
        formatNumber,
        debitMonth
      );
  }

  const doc = new jsPDF();
  const pageWidth = doc.internal.pageSize.getWidth();
  const yOffset = 25; // Increased offset to move content down for more top margin

  const customerState = extractStateFromAddress(data.customer.address);

  let currentYPosition = yOffset; // Initialize currentYPosition for consistent tracking

  // Receiver Details (Customer) - Centered for all types
  doc.setFontSize(11); // Increased font size
  doc.setFont(undefined, "bold");
  doc.text(data.customer.customerName, pageWidth / 2, currentYPosition, {
    align: "center",
  });
  currentYPosition += 10; // Increased spacing between customer name and address

  doc.setFont(undefined, "bold"); // Set font to bold for address
  let customerAddressLines = doc.splitTextToSize(data.customer.address, 150); // Increased width for centering
  let addressWithLabel: string[];

  if (customerAddressLines.length > 0) {
    addressWithLabel = [
      `ADDRESS: ${customerAddressLines[0]}`,
      ...customerAddressLines.slice(1),
    ];
  } else {
    addressWithLabel = ["ADDRESS:"];
  }

  addressWithLabel.forEach((line) => {
    doc.text(line, pageWidth / 2, currentYPosition, {
      align: "center",
    });
    currentYPosition += 4.5; // Increased spacing for each line
  });
  // GSTIN, PAN, MOB - Centered and condensed
  doc.setFont(undefined, "bold");
  const mobileText = data.customer.mobile
    ? `   MOB: ${data.customer.mobile}`
    : "";
  const gstinPanMobText = `GSTIN: ${data.customer.gstin}   PAN: ${data.customer.pan}${mobileText}`;
  doc.text(gstinPanMobText, pageWidth / 2, currentYPosition + 2, {
    align: "center",
  });
  currentYPosition += 15; // Increased spacing between customer details and "To:"

  // To: Section (Sender Details - Jubilant) - Left Aligned for all types
  doc.setFontSize(11); // Increased font size
  doc.setFont(undefined, "bold");
  doc.text("To:", 20, currentYPosition);
  currentYPosition += 7; // Increased spacing

  // Fixed Jubilant address for all customers
  const fixedJubilantAddress = [
    "The Sales Head",
    "Jubilant Agri & Consumer Products Ltd.",
    "Plot No 142, Chimes, 3rd Floor, Sector 44,",
    "Gurugram 3rd Floor, Sector -44, Gurugram,",
    "Haryana-122003",
  ];

  doc.setFont(undefined, "normal");
  fixedJubilantAddress.forEach((line) => {
    doc.text(line, 20, currentYPosition);
    currentYPosition += 5;
  });

  // Note: No GSTIN shown for the fixed address as per user requirement
  currentYPosition += 2; // Small spacing after address

  // Subject
  doc.setFontSize(11); // Increased font size
  doc.setFont(undefined, "normal");
  let subjectText = "";
  if (invoiceType === "godown") {
    subjectText = "Subject: Reimbursement of Rent Expenses";
  } else if (invoiceType === "main") {
    subjectText = "Subject: Reimbursement of Handling Expenses.";
  } else if (invoiceType === "freight") {
    subjectText = "Subject: Reimbursement of Freight.";
  }
  doc.text(subjectText, 20, currentYPosition + 7); // Increased spacing
  currentYPosition += 16; // Increased spacing

  // Dear Sir
  doc.setFontSize(11); // Increased font size
  doc.setFont(undefined, "normal");
  doc.text("Dear Sir", 20, currentYPosition);
  currentYPosition += 12; // Increased spacing

  // Introductory Text
  let introductoryText = "";
  if (invoiceType === "godown") {
    introductoryText =
     `Request to kindly reimburse the Godown Rent Expenses for the month of ${debitMonth} as per\ndetails given as under:`;
  } else if (invoiceType === "main") {
    introductoryText =
      `Request to kindly reimburse the Handling Expenses for the month of ${debitMonth} as per\ndetails given as under:`;
  } else if (invoiceType === "freight") {
    introductoryText =
       `Request to kindly reimburse the Freight Expenses for the month of ${debitMonth} as per details\ngiven as below:`;
  }
  doc.text(introductoryText, 20, currentYPosition);
  currentYPosition += 12; // Increased spacing

  let serviceRows: any[] = [];
  let totalAmountBeforeTax = 0;

  switch (invoiceType) {
    case "godown":
      serviceRows.push([
        "REIMBURSEMENT OF RENT EXPENSES", // Particulars
        formatNumber(data.quantityLifted), // Quantity as per data
        formatNumber(100), // Rate/MT (dynamic)
        formatNumber(data.godownRent), // Amt. In Ru. (dynamic)
      ]);
      totalAmountBeforeTax = data.godownRent;
      break;
    case "main":
      const mainParticularsText =
        "REIMBURSEMENT OF LOADING, UNLOADING & LOCAL TRANSPORTATION";
      const wrappedMainParticulars = doc.splitTextToSize(
        mainParticularsText,
        85
      ); // Wrap text to 85 units width

      // Push the first line with quantity and amount (no rate)
      serviceRows.push([
        wrappedMainParticulars[0], // First line of particulars
        formatNumber(data.quantityLifted), // Quantity as per data
        formatNumber(
          data.loadingCharges + data.unloadingCharges + data.localTransportation
        ), // Amount (no rate column)
      ]);

      // Push subsequent lines of particulars without other data
      for (let i = 1; i < wrappedMainParticulars.length; i++) {
        serviceRows.push([wrappedMainParticulars[i], "", ""]); // Empty strings for quantity and amount
      }

      totalAmountBeforeTax =
        data.loadingCharges + data.unloadingCharges + data.localTransportation;
      break;
    case "freight":
      serviceRows.push([
        "REIMBURSEMENT OF FREIGHT EXPENSES", // Particulars
        formatNumber(data.freightBalance), // Amount only (no quantity, no rate)
      ]);
      totalAmountBeforeTax = data.freightBalance;
      break;
  }

  let totalAmount = totalAmountBeforeTax;

  // Manual positioning of data
  currentYPosition += 8; // Increased space before data section

  doc.setFontSize(10); // Increased font size for data section
  doc.setFont(undefined, "bold");

  // Headers - Different layout based on invoice type
  if (invoiceType === "freight") {
    // Freight: Only Particulars and Amount
    doc.text("Particulars", 20, currentYPosition);
    doc.text("Amt. In Ru.", 180, currentYPosition, { align: "right" });
  } else if (invoiceType === "main") {
    // Main: Particulars, Quantity, Amount (no Rate)
    doc.text("Particulars", 20, currentYPosition);
    doc.text("Qty in MT", 110, currentYPosition, { align: "center" });
    doc.text("Amt. In Ru.", 180, currentYPosition, { align: "right" });
  } else {
    // Godown: Keep all columns as before
    doc.text("Particulars", 20, currentYPosition);
    doc.text("Rate/MT", 140, currentYPosition, { align: "center" });
    doc.text("Amt. In Ru.", 180, currentYPosition, { align: "right" });
  }
  currentYPosition += 6; // Increased spacing

  doc.setFont(undefined, "normal");

  // Service Rows - Different rendering based on invoice type
  serviceRows.forEach((row) => {
    doc.text(row[0], 20, currentYPosition); // Always print particulars

    // Only print other columns if they are not empty (i.e., it's the first line of a wrapped item)
    if (row[1] !== "") {
      if (invoiceType === "freight") {
        // Freight: Only show amount (no quantity, no rate)
        doc.text(row[1], 180, currentYPosition, { align: "right" }); // Amount only
      } else if (invoiceType === "main") {
        // Main: Show quantity and amount (no rate)
        doc.text(row[1], 110, currentYPosition, { align: "center" }); // Quantity restored
        doc.text(row[2], 180, currentYPosition, { align: "right" }); // Amount
      } else {
        // Godown: Show all columns
        doc.text(row[1], 110, currentYPosition, { align: "center" }); // Quantity
        doc.text(row[2], 140, currentYPosition, { align: "center" }); // Rate
        doc.text(row[3], 180, currentYPosition, { align: "right" }); // Amount
      }
    }
    currentYPosition += 6; // Increased spacing
  });

  // Total Row
  currentYPosition += 4; // Increased gap before total
  doc.setFont(undefined, "bold");
  doc.text("Total", 20, currentYPosition);
  doc.text(formatNumber(totalAmount), 180, currentYPosition, {
    align: "right",
  });
  currentYPosition += 6; // Increased spacing

  // Footer Section
  const finalY = currentYPosition + 10; // Increased space before footer
  doc.setFontSize(10); // Increased font size for footer
  doc.setFont(undefined, "normal");
  doc.text("Thanking You", 20, finalY);
  doc.text("Yours Faithfully", 20, finalY + 6); // Increased spacing
  doc.text("Authorised Signatory", 20, finalY + 18); // Increased spacing

  return doc.output("blob");
};
