import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';

type EntryType = 'Receipt' | 'Payment' | 'Contra' | 'Journal';

type PaymentMode = 'Cash' | 'Bank' | 'UPI' | 'Card' | 'Cheque';

interface DaybookEntry {
  id: number;

  date: string;

  time: string;

  voucherNo: string;

  type: EntryType;

  particulars: string;

  reference: string;

  paymentMode: PaymentMode;

  debit: number;

  credit: number;

  narration: string;

  debitAccount: string;

  creditAccount: string;
}

@Component({
  selector: 'app-daybook',

  standalone: true,

  imports: [CommonModule, FormsModule],

  templateUrl: './daybook.html',

  styleUrl: './daybook.scss',
})
export class Daybook {
  showDeleteConfirm = false;
  deleteEntryId: number | null = null;
  deleteVoucherNo = '';
  // =========================================================
  // FILTERS
  // =========================================================

  selectedDate = '2026-09-29';

  selectedEntryType: 'All' | EntryType = 'All';

  selectedPaymentMode: 'All' | PaymentMode = 'All';

  searchText = '';

  // =========================================================
  // BALANCE
  // =========================================================

  openingBalance = 125000;

  // =========================================================
  // MODAL
  // =========================================================

  showEntryForm = false;

  isEditMode = false;

  editingEntryId: number | null = null;

  newEntryType: EntryType = 'Receipt';

  // =========================================================
  // NEW / EDIT ENTRY FORM
  // =========================================================

  newEntry = {
    date: '2026-09-29',

    time: '',

    particulars: '',

    reference: '',

    paymentMode: 'Cash' as PaymentMode,

    debitAccount: '',

    creditAccount: '',

    amount: 0,

    narration: '',
  };

  // =========================================================
  // ACCOUNT MASTER
  // =========================================================

  cashBankAccounts = ['Cash', 'HDFC Bank', 'ICICI Bank', 'SBI Bank', 'Petty Cash'];

  receiptAccounts = ['Customer', 'Sales', 'Other Income', 'Advance from Customer'];

  paymentAccounts = [
    'Purchase',

    'Office Expense',

    'Salary Expense',

    'Electricity Expense',

    'Rent Expense',

    'Supplier',
  ];

  journalAccounts = [
    'Depreciation Expense',

    'Salary Expense',

    'Outstanding Expense',

    'Prepaid Expense',

    'Capital Account',

    'General Expense',
  ];

  // =========================================================
  // DAYBOOK DATA
  // =========================================================

  entries: DaybookEntry[] = [
    {
      id: 1,

      date: '29-09-2026',

      time: '09:15 AM',

      voucherNo: 'REC-0001',

      type: 'Receipt',

      particulars: 'Cash Sales',

      reference: 'INV-1001',

      paymentMode: 'Cash',

      debit: 0,

      credit: 15000,

      narration: 'Cash received against sales',

      debitAccount: 'Cash',

      creditAccount: 'Sales',
    },

    {
      id: 2,

      date: '29-09-2026',

      time: '10:30 AM',

      voucherNo: 'PAY-0001',

      type: 'Payment',

      particulars: 'Purchase Payment',

      reference: 'PUR-2045',

      paymentMode: 'Bank',

      debit: 8500,

      credit: 0,

      narration: 'Payment made to supplier',

      debitAccount: 'Purchase',

      creditAccount: 'HDFC Bank',
    },

    {
      id: 3,

      date: '29-09-2026',

      time: '11:45 AM',

      voucherNo: 'PAY-0002',

      type: 'Payment',

      particulars: 'Office Expense',

      reference: 'EXP-0012',

      paymentMode: 'Cash',

      debit: 5000,

      credit: 0,

      narration: 'Office miscellaneous expense',

      debitAccount: 'Office Expense',

      creditAccount: 'Cash',
    },

    {
      id: 4,

      date: '29-09-2026',

      time: '12:30 PM',

      voucherNo: 'REC-0002',

      type: 'Receipt',

      particulars: 'Customer Payment',

      reference: 'CUS-1020',

      paymentMode: 'UPI',

      debit: 0,

      credit: 22000,

      narration: 'Payment received from customer',

      debitAccount: 'Cash',

      creditAccount: 'Customer',
    },

    {
      id: 5,

      date: '29-09-2026',

      time: '01:15 PM',

      voucherNo: 'CON-0001',

      type: 'Contra',

      particulars: 'Cash Deposited to Bank',

      reference: 'CNT-001',

      paymentMode: 'Bank',

      debit: 10000,

      credit: 0,

      narration: 'Cash transferred to current account',

      debitAccount: 'HDFC Bank',

      creditAccount: 'Cash',
    },

    {
      id: 6,

      date: '29-09-2026',

      time: '02:00 PM',

      voucherNo: 'CON-0002',

      type: 'Contra',

      particulars: 'Bank Withdrawal',

      reference: 'CNT-002',

      paymentMode: 'Cash',

      debit: 0,

      credit: 7500,

      narration: 'Cash withdrawn from bank',

      debitAccount: 'Cash',

      creditAccount: 'HDFC Bank',
    },

    {
      id: 7,

      date: '29-09-2026',

      time: '03:10 PM',

      voucherNo: 'JRN-0001',

      type: 'Journal',

      particulars: 'Depreciation Expense',

      reference: 'JV-001',

      paymentMode: 'Bank',

      debit: 3000,

      credit: 0,

      narration: 'Monthly depreciation adjustment',

      debitAccount: 'Depreciation Expense',

      creditAccount: 'Capital Account',
    },

    {
      id: 8,

      date: '29-09-2026',

      time: '04:20 PM',

      voucherNo: 'REC-0003',

      type: 'Receipt',

      particulars: 'Customer Advance',

      reference: 'CUS-1050',

      paymentMode: 'Bank',

      debit: 0,

      credit: 18000,

      narration: 'Advance received from customer',

      debitAccount: 'HDFC Bank',

      creditAccount: 'Advance from Customer',
    },
  ];

  // =========================================================
  // FILTERED ENTRIES
  // =========================================================

  get filteredEntries(): DaybookEntry[] {
    const search = this.searchText.trim().toLowerCase();

    return this.entries.filter((entry) => {
      // -----------------------------------------
      // DATE
      // -----------------------------------------

      const entryDateParts = entry.date.split('-');

      const entryDate = `${entryDateParts[2]}-${entryDateParts[1]}-${entryDateParts[0]}`;

      const dateMatch = !this.selectedDate || entryDate === this.selectedDate;

      // -----------------------------------------
      // ENTRY TYPE
      // -----------------------------------------

      const typeMatch = this.selectedEntryType === 'All' || entry.type === this.selectedEntryType;

      // -----------------------------------------
      // PAYMENT MODE
      // -----------------------------------------

      const paymentModeMatch =
        this.selectedPaymentMode === 'All' || entry.paymentMode === this.selectedPaymentMode;

      // -----------------------------------------
      // SEARCH
      // -----------------------------------------

      const searchMatch =
        !search ||
        entry.voucherNo.toLowerCase().includes(search) ||
        entry.particulars.toLowerCase().includes(search) ||
        entry.reference.toLowerCase().includes(search) ||
        entry.narration.toLowerCase().includes(search) ||
        entry.paymentMode.toLowerCase().includes(search);

      return dateMatch && typeMatch && paymentModeMatch && searchMatch;
    });
  }

  // =========================================================
  // TOTAL DEBIT
  // =========================================================

  get totalDebit(): number {
    return this.filteredEntries.reduce(
      (total, entry) => total + Number(entry.debit),

      0,
    );
  }

  // =========================================================
  // TOTAL CREDIT
  // =========================================================

  get totalCredit(): number {
    return this.filteredEntries.reduce(
      (total, entry) => total + Number(entry.credit),

      0,
    );
  }

  // =========================================================
  // CLOSING BALANCE
  // =========================================================

  get closingBalance(): number {
    return Number(this.openingBalance) + this.totalCredit - this.totalDebit;
  }

  // =========================================================
  // TRANSACTION COUNT
  // =========================================================

  get transactionCount(): number {
    return this.filteredEntries.length;
  }

  // =========================================================
  // CURRENCY
  // =========================================================

  formatCurrency(amount: number): string {
    return new Intl.NumberFormat(
      'en-IN',

      {
        style: 'currency',

        currency: 'INR',

        minimumFractionDigits: 2,

        maximumFractionDigits: 2,
      },
    ).format(amount);
  }

  // =========================================================
  // RESET FILTERS
  // =========================================================

  resetFilters(): void {
    this.selectedDate = '';
    this.selectedEntryType = 'All';
    this.selectedPaymentMode = 'All';
    this.searchText = '';
  }

  // =========================================================
  // REFRESH
  // =========================================================

  refreshEntries(): void {
    // Local Angular data does not require
    // an API refresh.

    this.entries = [...this.entries];
  }

  // =========================================================
  // EXPORT
  // =========================================================

  exportEntries(): void {
    alert('Export functionality will be added in the next step.');
  }

  // =========================================================
  // OPEN NEW ENTRY
  // =========================================================

  openEntryForm(): void {
    this.showEntryForm = true;

    this.isEditMode = false;

    this.editingEntryId = null;

    this.newEntryType = 'Receipt';

    this.newEntry = {
      date: this.selectedDate || this.getTodayDate(),

      time: '',

      particulars: '',

      reference: '',

      paymentMode: 'Cash',

      debitAccount: '',

      creditAccount: '',

      amount: 0,

      narration: '',
    };
  }

  // =========================================================
  // EDIT ENTRY
  // =========================================================

  editEntry(entry: DaybookEntry): void {
    this.showEntryForm = true;

    this.isEditMode = true;

    this.editingEntryId = entry.id;

    this.newEntryType = entry.type;

    // Convert DD-MM-YYYY
    // to YYYY-MM-DD

    const dateParts = entry.date.split('-');

    const formDate = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;

    this.newEntry = {
      date: formDate,

      time: this.convertToInputTime(entry.time),

      particulars: entry.particulars,

      reference: entry.reference,

      paymentMode: entry.paymentMode,

      debitAccount: entry.debitAccount,

      creditAccount: entry.creditAccount,

      amount: entry.debit > 0 ? entry.debit : entry.credit,

      narration: entry.narration,
    };
  }

  // =========================================================
  // CONVERT DISPLAY TIME TO INPUT TIME
  // =========================================================

  convertToInputTime(time: string): string {
    if (!time) {
      return '';
    }

    const parts = time.trim().split(' ');

    if (parts.length !== 2) {
      return '';
    }

    const timeParts = parts[0].split(':');

    if (timeParts.length !== 2) {
      return '';
    }

    let hours = Number(timeParts[0]);

    const minutes = timeParts[1];

    const period = parts[1].toUpperCase();

    if (period === 'PM' && hours !== 12) {
      hours += 12;
    }

    if (period === 'AM' && hours === 12) {
      hours = 0;
    }

    return String(hours).padStart(2, '0') + ':' + minutes;
  }

  // =========================================================
  // CLOSE MODAL
  // =========================================================

  closeEntryForm(): void {
    this.showEntryForm = false;

    this.isEditMode = false;

    this.editingEntryId = null;
  }

  // =========================================================
  // BACKDROP CLICK
  // =========================================================

  onBackdropClick(event: MouseEvent): void {
    if (event.target === event.currentTarget) {
      this.closeEntryForm();
    }
  }

  // =========================================================
  // DEBIT ACCOUNTS
  // =========================================================

  getDebitAccounts(): string[] {
    switch (this.newEntryType) {
      case 'Receipt':
        return this.cashBankAccounts;

      case 'Payment':
        return this.paymentAccounts;

      case 'Contra':
        return this.cashBankAccounts;

      case 'Journal':
        return this.journalAccounts;

      default:
        return [];
    }
  }

  // =========================================================
  // CREDIT ACCOUNTS
  // =========================================================

  getCreditAccounts(): string[] {
    switch (this.newEntryType) {
      case 'Receipt':
        return this.receiptAccounts;

      case 'Payment':
        return this.cashBankAccounts;

      case 'Contra':
        return this.cashBankAccounts;

      case 'Journal':
        return this.journalAccounts;

      default:
        return [];
    }
  }

  // =========================================================
  // SAVE / UPDATE ENTRY
  // =========================================================

  saveEntry(): void {
    // -----------------------------------------
    // VALIDATION
    // -----------------------------------------

    if (!this.newEntry.date) {
      alert('Please select a date.');

      return;
    }

    if (!this.newEntry.particulars.trim()) {
      alert('Please enter particulars.');

      return;
    }

    if (!this.newEntry.amount || Number(this.newEntry.amount) <= 0) {
      alert('Please enter a valid amount.');

      return;
    }

    if (!this.newEntry.paymentMode) {
      alert('Please select payment mode.');

      return;
    }

    // -----------------------------------------
    // EDIT EXISTING ENTRY
    // -----------------------------------------

    if (this.isEditMode && this.editingEntryId !== null) {
      const index = this.entries.findIndex((entry) => entry.id === this.editingEntryId);

      if (index !== -1) {
        const existingEntry = this.entries[index];

        let debit = 0;

        let credit = 0;

        if (this.newEntryType === 'Receipt') {
          credit = Number(this.newEntry.amount);
        } else if (this.newEntryType === 'Payment') {
          debit = Number(this.newEntry.amount);
        } else if (this.newEntryType === 'Contra') {
          debit = Number(this.newEntry.amount);
        } else if (this.newEntryType === 'Journal') {
          debit = Number(this.newEntry.amount);
        }

        const dateParts = this.newEntry.date.split('-');

        const formattedDate = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;

        const displayTime = this.getDisplayTime(this.newEntry.time);

        this.entries[index] = {
          ...existingEntry,

          date: formattedDate,

          time: displayTime,

          type: this.newEntryType,

          particulars: this.newEntry.particulars.trim(),

          reference: this.newEntry.reference.trim(),

          paymentMode: this.newEntry.paymentMode,

          debit,

          credit,

          narration: this.newEntry.narration.trim(),

          debitAccount: this.newEntry.debitAccount,

          creditAccount: this.newEntry.creditAccount,
        };

        this.entries = [...this.entries];
      }

      this.closeEntryForm();

      return;
    }

    // -----------------------------------------
    // NEW ENTRY
    // -----------------------------------------

    let prefix = '';

    switch (this.newEntryType) {
      case 'Receipt':
        prefix = 'REC';

        break;

      case 'Payment':
        prefix = 'PAY';

        break;

      case 'Contra':
        prefix = 'CON';

        break;

      case 'Journal':
        prefix = 'JRN';

        break;
    }

    // -----------------------------------------
    // NEXT VOUCHER NUMBER
    // -----------------------------------------

    const sameTypeEntries = this.entries.filter((entry) => entry.type === this.newEntryType);

    const existingNumbers = sameTypeEntries

      .map((entry) => {
        const parts = entry.voucherNo.split('-');

        return Number(parts[1]);
      })

      .filter((number) => !isNaN(number));

    const highestNumber = existingNumbers.length > 0 ? Math.max(...existingNumbers) : 0;

    const nextNumber = highestNumber + 1;

    const voucherNo = `${prefix}-${String(nextNumber).padStart(4, '0')}`;

    // -----------------------------------------
    // TIME
    // -----------------------------------------

    const displayTime = this.getDisplayTime(this.newEntry.time);

    // -----------------------------------------
    // DATE
    // -----------------------------------------

    const dateParts = this.newEntry.date.split('-');

    const formattedDate = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;

    // -----------------------------------------
    // DEBIT / CREDIT
    // -----------------------------------------

    let debit = 0;

    let credit = 0;

    if (this.newEntryType === 'Receipt') {
      credit = Number(this.newEntry.amount);
    } else if (this.newEntryType === 'Payment') {
      debit = Number(this.newEntry.amount);
    } else if (this.newEntryType === 'Contra') {
      debit = Number(this.newEntry.amount);
    } else if (this.newEntryType === 'Journal') {
      debit = Number(this.newEntry.amount);
    }

    // -----------------------------------------
    // CREATE ENTRY
    // -----------------------------------------

    const newEntry: DaybookEntry = {
      id: this.getNextId(),

      date: formattedDate,

      time: displayTime,

      voucherNo,

      type: this.newEntryType,

      particulars: this.newEntry.particulars.trim(),

      reference: this.newEntry.reference.trim(),

      paymentMode: this.newEntry.paymentMode,

      debit,

      credit,

      narration: this.newEntry.narration.trim(),

      debitAccount: this.newEntry.debitAccount,

      creditAccount: this.newEntry.creditAccount,
    };

    // -----------------------------------------
    // ADD
    // -----------------------------------------

    this.entries.push(newEntry);

    this.entries = [...this.entries];

    // -----------------------------------------
    // SELECT NEW ENTRY DATE
    // -----------------------------------------

    this.selectedDate = this.newEntry.date;

    // -----------------------------------------
    // CLOSE
    // -----------------------------------------

    this.closeEntryForm();
  }

  // =========================================================
  // DELETE ENTRY
  // =========================================================

  // =========================================================
  // DELETE ENTRY
  // =========================================================

  deleteEntry(id: number): void {
    const entry = this.entries.find((item) => item.id === id);

    if (!entry) {
      console.error('Entry not found:', id);
      return;
    }

    // Store selected entry for confirmation
    this.deleteEntryId = id;
    this.deleteVoucherNo = entry.voucherNo;

    // Open confirmation modal
    this.showDeleteConfirm = true;
  }

  // =========================================================
  // CONFIRM DELETE
  // =========================================================

  confirmDelete(): void {
    if (this.deleteEntryId === null) {
      return;
    }

    const id = this.deleteEntryId;

    // Delete record
    this.entries = this.entries.filter((item) => item.id !== id);

    // Reset confirmation modal
    this.showDeleteConfirm = false;
    this.deleteEntryId = null;
    this.deleteVoucherNo = '';

    // Update localStorage
    localStorage.setItem('daybookEntries', JSON.stringify(this.entries));
  }

  // =========================================================
  // CANCEL DELETE
  // =========================================================

  cancelDelete(): void {
    this.showDeleteConfirm = false;
    this.deleteEntryId = null;
    this.deleteVoucherNo = '';
  }

  // =========================================================
  // NEXT ID
  // =========================================================

  getNextId(): number {
    if (this.entries.length === 0) {
      return 1;
    }

    return Math.max(...this.entries.map((entry) => entry.id)) + 1;
  }

  // =========================================================
  // DISPLAY TIME
  // =========================================================

  getDisplayTime(time: string): string {
    if (!time) {
      const now = new Date();

      return now.toLocaleTimeString(
        'en-IN',

        {
          hour: '2-digit',

          minute: '2-digit',

          hour12: true,
        },
      );
    }

    const parts = time.split(':');

    if (parts.length !== 2) {
      return time;
    }

    const hours = Number(parts[0]);

    const minutes = parts[1];

    const suffix = hours >= 12 ? 'PM' : 'AM';

    const displayHour = hours % 12 || 12;

    return `${String(displayHour).padStart(2, '0')}` + `:${minutes} ${suffix}`;
  }

  // =========================================================
  // TODAY
  // =========================================================

  getTodayDate(): string {
    const today = new Date();

    const year = today.getFullYear();

    const month = String(today.getMonth() + 1).padStart(2, '0');

    const day = String(today.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  }

  exportExcel(): void {
    const data = this.filteredEntries.map((entry, index) => ({
      'S.No': index + 1,
      Date: entry.date,
      Time: entry.time,
      'Voucher No.': entry.voucherNo,
      'Entry Type': entry.type,
      Particulars: entry.particulars,
      Reference: entry.reference,
      'Payment Mode': entry.paymentMode,
      Debit: entry.debit,
      Credit: entry.credit,
      Narration: entry.narration,
    }));

    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    const workbook: XLSX.WorkBook = {
      Sheets: {
        Daybook: worksheet,
      },
      SheetNames: ['Daybook'],
    };

    XLSX.writeFile(workbook, `Daybook_Report.xlsx`);
  }

  exportPdf(): void {
    const doc = new jsPDF('l', 'mm', 'a4');

    // -----------------------------------------
    // TITLE
    // -----------------------------------------

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(18);
    doc.text('Daybook Report', 14, 15);

    // -----------------------------------------
    // REPORT DETAILS
    // -----------------------------------------

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);

    doc.text(`Date: ${this.selectedDate || 'All Dates'}`, 14, 22);

    doc.text(`Transactions: ${this.transactionCount}`, 80, 22);

    doc.text(`Opening Balance: ${this.formatPdfCurrency(this.openingBalance)}`, 145, 22);

    // -----------------------------------------
    // TABLE DATA
    // -----------------------------------------

    const tableData = this.filteredEntries.map((entry, index) => [
      index + 1,
      entry.date,
      entry.time,
      entry.voucherNo,
      entry.type,
      entry.particulars,
      entry.reference || '-',
      entry.paymentMode,

      entry.debit > 0 ? this.formatPdfCurrency(entry.debit) : '-',

      entry.credit > 0 ? this.formatPdfCurrency(entry.credit) : '-',

      entry.narration || '-',
    ]);

    // -----------------------------------------
    // TABLE
    // -----------------------------------------

    autoTable(doc, {
      startY: 28,

      head: [
        [
          '#',
          'Date',
          'Time',
          'Voucher No.',
          'Type',
          'Particulars',
          'Reference',
          'Payment Mode',
          'Debit',
          'Credit',
          'Narration',
        ],
      ],

      body: tableData,

      theme: 'grid',

      styles: {
        font: 'helvetica',
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
      },

      headStyles: {
        font: 'helvetica',
        fontStyle: 'bold',
        fontSize: 7,
      },

      columnStyles: {
        0: {
          cellWidth: 8,
          halign: 'center',
        },

        1: {
          cellWidth: 22,
        },

        2: {
          cellWidth: 20,
        },

        3: {
          cellWidth: 25,
        },

        4: {
          cellWidth: 20,
        },

        5: {
          cellWidth: 35,
        },

        6: {
          cellWidth: 25,
        },

        7: {
          cellWidth: 23,
        },

        8: {
          cellWidth: 27,
          halign: 'right',
        },

        9: {
          cellWidth: 27,
          halign: 'right',
        },

        10: {
          cellWidth: 45,
        },
      },
    });

    // -----------------------------------------
    // TOTALS
    // -----------------------------------------

    const finalY = (doc as any).lastAutoTable.finalY + 8;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);

    doc.text(`Total Debit: ${this.formatPdfCurrency(this.totalDebit)}`, 14, finalY);

    doc.text(`Total Credit: ${this.formatPdfCurrency(this.totalCredit)}`, 95, finalY);

    doc.text(`Closing Balance: ${this.formatPdfCurrency(this.closingBalance)}`, 180, finalY);

    // -----------------------------------------
    // SAVE PDF
    // -----------------------------------------

    doc.save('Daybook_Report.pdf');
  }

  formatPdfCurrency(amount: number): string {
    return `Rs. ${amount.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  }
}
