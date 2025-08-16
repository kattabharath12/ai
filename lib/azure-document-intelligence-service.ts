

import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { readFile } from "fs/promises";

export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

export interface ExtractedFieldData {
  [key: string]: string | number;
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: AzureDocumentIntelligenceConfig;

  constructor(config: AzureDocumentIntelligenceConfig) {
    this.config = config;
    this.client = new DocumentAnalysisClient(
      this.config.endpoint,
      new AzureKeyCredential(this.config.apiKey)
    );
  }

  async extractDataFromDocument(
    documentPathOrBuffer: string | Buffer,
    documentType: string
  ): Promise<ExtractedFieldData> {
    try {
      console.log('🔍 [Azure DI] Processing document with Azure Document Intelligence...');
      
      // Get document buffer - either from file path or use provided buffer
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Determine the model to use based on document type
      const modelId = this.getModelIdForDocumentType(documentType);
      console.log('🔍 [Azure DI] Using model:', modelId);
      
      // Analyze the document
      const poller = await this.client.beginAnalyzeDocument(modelId, documentBuffer);
      const result = await poller.pollUntilDone();
      
      console.log('✅ [Azure DI] Document analysis completed');
      
      // Extract the data based on document type
      return this.extractTaxDocumentFields(result, documentType);
    } catch (error: any) {
      console.error('❌ [Azure DI] Processing error:', error);
      throw new Error(`Azure Document Intelligence processing failed: ${error?.message || 'Unknown error'}`);
    }
  }

  private getModelIdForDocumentType(documentType: string): string {
    switch (documentType) {
      case 'W2':
        return 'prebuilt-tax.us.w2';
      case 'FORM_1099_INT':
        return 'prebuilt-tax.us.1099int';
      case 'FORM_1099_DIV':
        return 'prebuilt-tax.us.1099div';
      case 'FORM_1099_MISC':
        return 'prebuilt-tax.us.1099misc';
      case 'FORM_1099_NEC':
        return 'prebuilt-tax.us.1099nec';
      default:
        // Use general document model for other types
        return 'prebuilt-document';
    }
  }

  private extractTaxDocumentFields(result: any, documentType: string): ExtractedFieldData {
    const extractedData: ExtractedFieldData = {};
    
    // Extract text content
    extractedData.fullText = result.content || '';
    
    // Extract form fields
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      if (document.fields) {
        // Process fields based on document type
        switch (documentType) {
          case 'W2':
            return this.processW2Fields(document.fields, extractedData);
          case 'FORM_1099_INT':
            return this.process1099IntFields(document.fields, extractedData);
          case 'FORM_1099_DIV':
            return this.process1099DivFields(document.fields, extractedData);
          case 'FORM_1099_MISC':
            return this.process1099MiscFields(document.fields, extractedData);
          case 'FORM_1099_NEC':
            return this.process1099NecFields(document.fields, extractedData);
          default:
            return this.processGenericFields(document.fields, extractedData);
        }
      }
    }
    
    // Extract key-value pairs from tables if available
    if (result.keyValuePairs) {
      for (const kvp of result.keyValuePairs) {
        const key = kvp.key?.content?.trim();
        const value = kvp.value?.content?.trim();
        if (key && value) {
          extractedData[key] = value;
        }
      }
    }
    
    return extractedData;
  }

  private processW2Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const w2Data = { ...baseData };
    
    // W2 specific field mappings
    const w2FieldMappings = {
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN',
      'Employee.Address': 'employeeAddress',
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'WagesAndTips': 'wages',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'SocialSecurityWages': 'socialSecurityWages',
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'MedicareWagesAndTips': 'medicareWages',
      'MedicareTaxWithheld': 'medicareTaxWithheld',
      'SocialSecurityTips': 'socialSecurityTips',
      'AllocatedTips': 'allocatedTips',
      'StateWagesTipsEtc': 'stateWages',
      'StateIncomeTax': 'stateTaxWithheld',
      'LocalWagesTipsEtc': 'localWages',
      'LocalIncomeTax': 'localTaxWithheld'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(w2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        w2Data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    // Enhanced personal info extraction with better fallback handling
    console.log('🔍 [Azure DI] Extracting personal information from W2...');
    
    // Employee Name - try multiple field variations
    if (!w2Data.employeeName) {
      const nameFields = ['Employee.Name', 'EmployeeName', 'Employee_Name', 'RecipientName'];
      for (const fieldName of nameFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeName = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee name:', w2Data.employeeName);
          break;
        }
      }
    }
    
    // Employee SSN - try multiple field variations
    if (!w2Data.employeeSSN) {
      const ssnFields = ['Employee.SSN', 'EmployeeSSN', 'Employee_SSN', 'RecipientTIN'];
      for (const fieldName of ssnFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeSSN = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee SSN:', w2Data.employeeSSN);
          break;
        }
      }
    }
    
    // Employee Address - try multiple field variations
    if (!w2Data.employeeAddress) {
      const addressFields = ['Employee.Address', 'EmployeeAddress', 'Employee_Address', 'RecipientAddress'];
      for (const fieldName of addressFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeAddress = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee address:', w2Data.employeeAddress);
          break;
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!w2Data.employeeName || !w2Data.employeeSSN || !w2Data.employeeAddress) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some personal info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!w2Data.employeeName && personalInfoFromOCR.name) {
        w2Data.employeeName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted employee name from OCR:', w2Data.employeeName);
      }
      
      if (!w2Data.employeeSSN && personalInfoFromOCR.ssn) {
        w2Data.employeeSSN = personalInfoFromOCR.ssn;
        console.log('✅ [Azure DI] Extracted employee SSN from OCR:', w2Data.employeeSSN);
      }
      
      if (!w2Data.employeeAddress && personalInfoFromOCR.address) {
        w2Data.employeeAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted employee address from OCR:', w2Data.employeeAddress);
      }
    }
    
    // OCR fallback for Box 1 wages if not found in structured fields
    if (!w2Data.wages && baseData.fullText) {
      console.log('🔍 [Azure DI] Wages not found in structured fields, attempting OCR extraction...');
      const wagesFromOCR = this.extractWagesFromOCR(baseData.fullText as string);
      if (wagesFromOCR > 0) {
        console.log('✅ [Azure DI] Successfully extracted wages from OCR:', wagesFromOCR);
        w2Data.wages = wagesFromOCR;
      }
    }
    
    return w2Data;
  }

  private process1099IntFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'InterestIncome': 'interestIncome',
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'InvestmentExpenses': 'investmentExpenses',
      'ForeignTaxPaid': 'foreignTaxPaid',
      'TaxExemptInterest': 'taxExemptInterest'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    return data;
  }

  private process1099DivFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'OrdinaryDividends': 'ordinaryDividends',
      'QualifiedDividends': 'qualifiedDividends',
      'TotalCapitalGainDistributions': 'totalCapitalGain',
      'NondividendDistributions': 'nondividendDistributions',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'Section199ADividends': 'section199ADividends'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    return data;
  }

  private process1099MiscFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'Rents': 'rents',
      'Royalties': 'royalties',
      'OtherIncome': 'otherIncome',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FishingBoatProceeds': 'fishingBoatProceeds',
      'MedicalAndHealthCarePayments': 'medicalHealthPayments',
      'NonemployeeCompensation': 'nonemployeeCompensation'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    return data;
  }

  private process1099NecFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    return data;
  }

  private processGenericFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    for (const [fieldName, field] of Object.entries(fields || {})) {
      if (field && typeof field === 'object' && 'value' in field && (field as any).value !== undefined) {
        data[fieldName] = (field as any).value;
      }
    }
    
    return data;
  }

  private parseAmount(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[$,\s]/g, '');
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }

  /**
   * Extracts personal information from OCR text using improved regex patterns
   * Specifically designed for W-2 form OCR text patterns
   */
  private extractPersonalInfoFromOCR(ocrText: string): {
    name?: string;
    ssn?: string;
    address?: string;
  } {
    if (process.env.NODE_ENV === 'development') {
      console.log('🔍 [Azure DI OCR] Searching for personal info in OCR text...');
    }
    
    const personalInfo: { name?: string; ssn?: string; address?: string } = {};
    
    // Extract employee name - improved patterns for W-2 format
    // Handles patterns like "e Employee's first name and initial Last name Michelle Hicks"
    const namePatterns = [
      // Pattern for "e Employee's first name and initial Last name Michelle Hicks"
      /e\s+Employee's\s+first\s+name\s+and\s+initial\s+Last\s+name\s+([A-Za-z\s]+?)(?:\n|Employee's\s+address|$)/i,
      // Pattern for "Employee's first name and initial Last name Michelle Hicks"
      /Employee's\s+first\s+name\s+and\s+initial\s+Last\s+name\s+([A-Za-z\s]+?)(?:\n|Employee's\s+address|$)/i,
      // Fallback pattern for simpler formats
      /Employee[:\s]+([A-Za-z\s]+?)(?:\n|Employee's\s+address|Employee's\s+social|SSN|Social|Address|$)/i,
      // Additional fallback for "Employee Name:" format
      /Employee\s+Name[:\s]+([A-Za-z\s]+?)(?:\n|Employee's\s+address|Employee's\s+social|SSN|Social|Address|$)/i
    ];
    
    for (const pattern of namePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.name = match[1].trim();
        if (process.env.NODE_ENV === 'development') {
          console.log('🔍 [Azure DI OCR] Found employee name:', personalInfo.name);
        }
        break;
      }
    }
    
    // Extract SSN - enhanced patterns for W-2 format (keeping existing working patterns)
    const ssnPatterns = [
      // W-2 specific pattern
      /Employee's\s+social\s+security\s+number\s*\n(\d{3}-\d{2}-\d{4})/i,
      /social\s+security\s+number\s*\n(\d{3}-\d{2}-\d{4})/i,
      // Existing working patterns
      /SSN[:\s]*(\d{3}-\d{2}-\d{4})/i,
      /Social\s+Security[:\s]*(\d{3}-\d{2}-\d{4})/i,
      /(\d{3}-\d{2}-\d{4})/
    ];
    
    for (const pattern of ssnPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.ssn = match[1];
        if (process.env.NODE_ENV === 'development') {
          console.log('🔍 [Azure DI OCR] Found employee SSN:', personalInfo.ssn);
        }
        break;
      }
    }
    
    // Extract address - improved patterns for W-2 format
    // Handles "Employee's address and ZIP code" followed by multi-line address
    const addressPatterns = [
      // Primary pattern for "Employee's address and ZIP code" followed by address lines
      /Employee's\s+address\s+and\s+ZIP\s+code\s*\n([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|Employee's\s+social|Employer|$)/i,
      // Alternative pattern with more flexible spacing and line breaks
      /Employee's\s+address[^\n]*\n([^\n]+(?:\n[0-9A-Za-z][^\n]*)*?)(?:\n\s*\n|Employee's\s+social|Employer|$)/i,
      // Fallback pattern for simpler address formats
      /address\s+and\s+ZIP\s+code[^\n]*\n([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|social\s+security|Employer|$)/i,
      // Generic address pattern as last resort
      /Address[:\s]+([^\n]+(?:\n[^\n]+)*?)(?:\n\n|Employee|Employer|$)/i
    ];
    
    for (const pattern of addressPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        // Clean up the address: normalize whitespace and join lines with spaces
        personalInfo.address = match[1].trim().replace(/\n/g, ' ').replace(/\s+/g, ' ');
        if (process.env.NODE_ENV === 'development') {
          console.log('🔍 [Azure DI OCR] Found employee address:', personalInfo.address);
        }
        break;
      }
    }
    
    return personalInfo;
  }

  /**
   * Extracts wages from OCR text using regex patterns for Box 1
   */
  private extractWagesFromOCR(ocrText: string): number {
    console.log('🔍 [Azure DI OCR] Searching for wages in OCR text...');
    
    // Multiple regex patterns to match Box 1 wages
    const wagePatterns = [
      // Pattern: "1 Wages, tips, other compensation 161130.48"
      /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation\s+([\d,]+\.?\d*)/i,
      // Pattern: "1. Wages, tips, other compensation: $161,130.48"
      /\b1\.?\s*Wages[,\s]*tips[,\s]*other\s+compensation[:\s]+\$?([\d,]+\.?\d*)/i,
      // Pattern: "Box 1 161130.48" or "1 161130.48"
      /\b(?:Box\s*)?1\s+\$?([\d,]+\.?\d*)/i,
      // Pattern: "Wages and tips 161130.48"
      /Wages\s+and\s+tips\s+\$?([\d,]+\.?\d*)/i,
      // Pattern: "1 Wages, tips, other compensation" followed by amount on next line
      /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation[\s\n]+\$?([\d,]+\.?\d*)/i
    ];

    for (const pattern of wagePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const wageString = match[1];
        console.log('🔍 [Azure DI OCR] Found wage match:', wageString, 'using pattern:', pattern.source);
        
        // Parse the amount
        const cleanedAmount = wageString.replace(/[,$\s]/g, '');
        const parsedAmount = parseFloat(cleanedAmount);
        
        if (!isNaN(parsedAmount) && parsedAmount > 0) {
          console.log('✅ [Azure DI OCR] Successfully parsed wages:', parsedAmount);
          return parsedAmount;
        }
      }
    }

    console.log('⚠️ [Azure DI OCR] No wages found in OCR text');
    return 0;
  }

  async processW2Document(documentPathOrBuffer: string | Buffer): Promise<ExtractedFieldData> {
    const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, 'W2');
    
    return {
      documentType: 'FORM_W2',
      employerName: extractedData.employerName || '',
      employerEIN: extractedData.employerEIN || '',
      employeeName: extractedData.employeeName || '',
      employeeSSN: extractedData.employeeSSN || '',
      employeeAddress: extractedData.employeeAddress || '',
      wages: this.parseAmount(extractedData.wages) || 0,
      federalTaxWithheld: this.parseAmount(extractedData.federalTaxWithheld) || 0,
      socialSecurityWages: this.parseAmount(extractedData.socialSecurityWages) || 0,
      medicareWages: this.parseAmount(extractedData.medicareWages) || 0,
      socialSecurityTaxWithheld: this.parseAmount(extractedData.socialSecurityTaxWithheld) || 0,
      medicareTaxWithheld: this.parseAmount(extractedData.medicareTaxWithheld) || 0,
      stateWages: this.parseAmount(extractedData.stateWages) || 0,
      stateTaxWithheld: this.parseAmount(extractedData.stateTaxWithheld) || 0,
      ...extractedData
    };
  }

  async process1099Document(documentPathOrBuffer: string | Buffer, documentType: string): Promise<ExtractedFieldData> {
    const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, documentType);
    
    return {
      documentType,
      payerName: extractedData.payerName || '',
      payerTIN: extractedData.payerTIN || '',
      recipientName: extractedData.recipientName || '',
      recipientTIN: extractedData.recipientTIN || '',
      interestIncome: this.parseAmount(extractedData.interestIncome) || 0,
      ordinaryDividends: this.parseAmount(extractedData.ordinaryDividends) || 0,
      nonemployeeCompensation: this.parseAmount(extractedData.nonemployeeCompensation) || 0,
      federalTaxWithheld: this.parseAmount(extractedData.federalTaxWithheld) || 0,
      ...extractedData
    };
  }
}

// Singleton instance
let azureDocumentIntelligenceService: AzureDocumentIntelligenceService | null = null;

export function getAzureDocumentIntelligenceService(): AzureDocumentIntelligenceService {
  if (!azureDocumentIntelligenceService) {
    const config: AzureDocumentIntelligenceConfig = {
      endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT!,
      apiKey: process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY!,
    };

    if (!config.endpoint || !config.apiKey) {
      throw new Error('Azure Document Intelligence configuration missing. Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_API_KEY environment variables.');
    }

    azureDocumentIntelligenceService = new AzureDocumentIntelligenceService(config);
  }

  return azureDocumentIntelligenceService;
}

export function createAzureDocumentIntelligenceConfig(): AzureDocumentIntelligenceConfig {
  return {
    endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT!,
    apiKey: process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY!,
  };
}

