
import { Form1040Data, W2ToForm1040Mapping } from './form-1040-types';

export class W2ToForm1040Mapper {
  /**
   * Maps W2 extracted data to Form 1040 fields
   */
  static mapW2ToForm1040(w2Data: any, existingForm1040?: Partial<Form1040Data>): Partial<Form1040Data> {
    const form1040Data: Partial<Form1040Data> = {
      ...existingForm1040,
    };

    // Personal Information Mapping
    if (w2Data.employeeName && !form1040Data.firstName) {
      const nameParts = w2Data.employeeName.trim().split(/\s+/);
      form1040Data.firstName = nameParts[0] || '';
      form1040Data.lastName = nameParts.slice(1).join(' ') || '';
    }

    if (w2Data.employeeSSN && !form1040Data.ssn) {
      form1040Data.ssn = this.formatSSN(w2Data.employeeSSN);
    }

    if (w2Data.employeeAddress && !form1040Data.address) {
      const addressParts = this.parseAddress(w2Data.employeeAddress);
      form1040Data.address = addressParts.street;
      form1040Data.city = addressParts.city;
      form1040Data.state = addressParts.state;
      form1040Data.zipCode = addressParts.zipCode;
    }

    // Income Mapping
    // Line 1: Total amount from Form(s) W-2, box 1
    const wages = this.parseAmount(w2Data.wages);
    if (wages > 0) {
      form1040Data.line1 = (form1040Data.line1 || 0) + wages;
    }

    // Line 25a: Federal income tax withheld from Form(s) W-2
    const federalTaxWithheld = this.parseAmount(w2Data.federalTaxWithheld);
    if (federalTaxWithheld > 0) {
      form1040Data.line25a = (form1040Data.line25a || 0) + federalTaxWithheld;
    }

    // Calculate total income (Line 9) - simplified calculation
    form1040Data.line9 = this.calculateTotalIncome(form1040Data);

    // Calculate AGI (Line 11) - simplified (no adjustments for now)
    form1040Data.line11 = form1040Data.line9;

    // Apply standard deduction (Line 12) if not itemizing
    if (!form1040Data.line12 && form1040Data.filingStatus) {
      form1040Data.line12 = this.getStandardDeduction(form1040Data.filingStatus);
    }

    // Calculate taxable income (Line 15)
    form1040Data.line15 = Math.max(0, (form1040Data.line11 || 0) - (form1040Data.line12 || 0) - (form1040Data.line13 || 0));

    // Calculate tax liability (Line 16) - simplified
    form1040Data.line16 = this.calculateTaxLiability(form1040Data.line15 || 0, form1040Data.filingStatus);

    // Calculate total tax (Line 24) - simplified
    form1040Data.line24 = (form1040Data.line16 || 0) + (form1040Data.line17 || 0) + (form1040Data.line23 || 0);

    // Calculate total payments (Line 32)
    form1040Data.line32 = (form1040Data.line25a || 0) + (form1040Data.line25b || 0) + (form1040Data.line25c || 0) + (form1040Data.line25d || 0);

    // Calculate refund or amount owed
    const totalTax = form1040Data.line24 || 0;
    const totalPayments = form1040Data.line32 || 0;

    if (totalPayments > totalTax) {
      // Refund
      form1040Data.line33 = totalPayments - totalTax;
      form1040Data.line34 = form1040Data.line33; // Default to full refund
      form1040Data.line37 = 0;
    } else {
      // Amount owed
      form1040Data.line33 = 0;
      form1040Data.line34 = 0;
      form1040Data.line37 = totalTax - totalPayments;
    }

    return form1040Data;
  }

  /**
   * Creates a mapping summary showing what W2 fields mapped to which 1040 lines
   */
  static createMappingSummary(w2Data: any): Array<{
    w2Field: string;
    w2Value: any;
    form1040Line: string;
    form1040Value: any;
    description: string;
  }> {
    const mappings = [];

    if (w2Data.employeeName) {
      mappings.push({
        w2Field: 'Employee Name',
        w2Value: w2Data.employeeName,
        form1040Line: 'Header',
        form1040Value: w2Data.employeeName,
        description: 'Taxpayer name from W2'
      });
    }

    if (w2Data.employeeSSN) {
      mappings.push({
        w2Field: 'Employee SSN',
        w2Value: w2Data.employeeSSN,
        form1040Line: 'Header',
        form1040Value: this.formatSSN(w2Data.employeeSSN),
        description: 'Taxpayer SSN from W2'
      });
    }

    if (w2Data.wages) {
      mappings.push({
        w2Field: 'Box 1 - Wages',
        w2Value: w2Data.wages,
        form1040Line: 'Line 1',
        form1040Value: this.parseAmount(w2Data.wages),
        description: 'Wages, tips, other compensation'
      });
    }

    if (w2Data.federalTaxWithheld) {
      mappings.push({
        w2Field: 'Box 2 - Federal Tax Withheld',
        w2Value: w2Data.federalTaxWithheld,
        form1040Line: 'Line 25a',
        form1040Value: this.parseAmount(w2Data.federalTaxWithheld),
        description: 'Federal income tax withheld'
      });
    }

    if (w2Data.socialSecurityWages) {
      mappings.push({
        w2Field: 'Box 3 - Social Security Wages',
        w2Value: w2Data.socialSecurityWages,
        form1040Line: 'Informational',
        form1040Value: this.parseAmount(w2Data.socialSecurityWages),
        description: 'Social security wages (informational)'
      });
    }

    if (w2Data.medicareWages) {
      mappings.push({
        w2Field: 'Box 5 - Medicare Wages',
        w2Value: w2Data.medicareWages,
        form1040Line: 'Informational',
        form1040Value: this.parseAmount(w2Data.medicareWages),
        description: 'Medicare wages and tips (informational)'
      });
    }

    return mappings;
  }

  private static parseAmount(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[$,\s]/g, '');
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }

  private static formatSSN(ssn: string): string {
    if (!ssn) return '';
    // Remove all non-digits
    const cleaned = ssn.replace(/\D/g, '');
    // Format as XXX-XX-XXXX
    if (cleaned.length === 9) {
      return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 5)}-${cleaned.slice(5)}`;
    }
    return cleaned;
  }

  private static parseAddress(address: string): {
    street: string;
    city: string;
    state: string;
    zipCode: string;
  } {
    // Simple address parsing - can be enhanced
    const parts = address.split(',').map(part => part.trim());
    
    if (parts.length >= 3) {
      const street = parts.slice(0, -2).join(', ');
      const city = parts[parts.length - 2];
      const stateZip = parts[parts.length - 1];
      const stateZipMatch = stateZip.match(/^([A-Z]{2})\s*(\d{5}(-\d{4})?)$/);
      
      return {
        street,
        city,
        state: stateZipMatch ? stateZipMatch[1] : '',
        zipCode: stateZipMatch ? stateZipMatch[2] : ''
      };
    }
    
    return {
      street: address,
      city: '',
      state: '',
      zipCode: ''
    };
  }

  private static calculateTotalIncome(form1040Data: Partial<Form1040Data>): number {
    return (
      (form1040Data.line1 || 0) +
      (form1040Data.line2b || 0) +
      (form1040Data.line3b || 0) +
      (form1040Data.line4b || 0) +
      (form1040Data.line5b || 0) +
      (form1040Data.line6b || 0) +
      (form1040Data.line7 || 0) +
      (form1040Data.line8 || 0)
    );
  }

  private static getStandardDeduction(filingStatus: any): number {
    const STANDARD_DEDUCTION_2023: Record<string, number> = {
      'SINGLE': 13850,
      'MARRIED_FILING_JOINTLY': 27700,
      'MARRIED_FILING_SEPARATELY': 13850,
      'HEAD_OF_HOUSEHOLD': 20800,
      'QUALIFYING_SURVIVING_SPOUSE': 27700
    };
    
    return STANDARD_DEDUCTION_2023[filingStatus] || STANDARD_DEDUCTION_2023['SINGLE'];
  }

  private static calculateTaxLiability(taxableIncome: number, filingStatus: any): number {
    // Simplified tax calculation using 2023 tax brackets
    const brackets: Array<{ min: number; max: number; rate: number }> = 
      filingStatus === 'MARRIED_FILING_JOINTLY' ? [
        { min: 0, max: 22000, rate: 0.10 },
        { min: 22000, max: 89450, rate: 0.12 },
        { min: 89450, max: 190750, rate: 0.22 },
        { min: 190750, max: 364200, rate: 0.24 },
        { min: 364200, max: 462500, rate: 0.32 },
        { min: 462500, max: 693750, rate: 0.35 },
        { min: 693750, max: Infinity, rate: 0.37 }
      ] : [
        { min: 0, max: 11000, rate: 0.10 },
        { min: 11000, max: 44725, rate: 0.12 },
        { min: 44725, max: 95375, rate: 0.22 },
        { min: 95375, max: 182050, rate: 0.24 },
        { min: 182050, max: 231250, rate: 0.32 },
        { min: 231250, max: 578125, rate: 0.35 },
        { min: 578125, max: Infinity, rate: 0.37 }
      ];

    let tax = 0;
    let remainingIncome = taxableIncome;

    for (const bracket of brackets) {
      if (remainingIncome <= 0) break;
      
      const taxableAtThisBracket = Math.min(remainingIncome, bracket.max - bracket.min);
      tax += taxableAtThisBracket * bracket.rate;
      remainingIncome -= taxableAtThisBracket;
    }

    return Math.round(tax * 100) / 100; // Round to 2 decimal places
  }

  /**
   * Validates that W2 data can be properly mapped to 1040
   */
  static validateW2DataForMapping(w2Data: any): {
    isValid: boolean;
    errors: string[];
    warnings: string[];
  } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Required fields
    if (!w2Data.wages && !w2Data.line1) {
      errors.push('W2 wages (Box 1) is required but not found');
    }

    if (!w2Data.employeeSSN && !w2Data.ssn) {
      errors.push('Employee SSN is required but not found');
    }

    if (!w2Data.employeeName && !w2Data.firstName && !w2Data.lastName) {
      errors.push('Employee name is required but not found');
    }

    // Warnings for missing optional but important fields
    if (!w2Data.federalTaxWithheld) {
      warnings.push('Federal tax withheld (Box 2) not found - no withholdings will be applied');
    }

    if (!w2Data.employerName) {
      warnings.push('Employer name not found - may be needed for verification');
    }

    if (!w2Data.employerEIN) {
      warnings.push('Employer EIN not found - may be needed for verification');
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }
}
