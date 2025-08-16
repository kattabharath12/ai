
import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { W2ToForm1040Mapper } from "@/lib/w2-to-1040-mapping";
import { Form1040Data } from "@/lib/form-1040-types";

export const dynamic = "force-dynamic";

// GET: Retrieve 1040 form data with W2 mappings
export async function GET(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 GET] Starting form 1040 data retrieval for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      console.log("❌ [1040 GET] No session found");
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      console.log("❌ [1040 GET] User not found");
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    // Get tax return with all related data
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      },
      include: {
        incomeEntries: true,
        deductionEntries: true,
        dependents: true,
        documents: {
          where: { 
            documentType: 'W2',
            processingStatus: 'COMPLETED'
          },
          include: {
            extractedEntries: true
          }
        }
      }
    });

    if (!taxReturn) {
      console.log("❌ [1040 GET] Tax return not found");
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    // Check if there's existing 1040 data
    let form1040Data: Partial<Form1040Data> = {};
    
    // If there's saved 1040 data in a separate table or JSON field, load it
    // For now, we'll construct it from the tax return data
    
    // Get W2 documents and their extracted data
    const w2Documents = taxReturn.documents.filter((doc: any) => doc.documentType === 'W2');
    const w2MappingData = [];

    // Process each W2 document and map to 1040 form
    console.log(`🔍 [1040 GET] Processing ${w2Documents.length} W2 documents`);
    
    for (const w2Doc of w2Documents) {
      console.log(`🔍 [1040 GET] Processing W2 document: ${w2Doc.fileName} (ID: ${w2Doc.id})`);
      console.log(`🔍 [1040 GET] W2 document extractedData:`, JSON.stringify(w2Doc.extractedData, null, 2));
      
      if (w2Doc.extractedData && typeof w2Doc.extractedData === 'object') {
        const extractedData = w2Doc.extractedData as any;
        
        // Try different data structure paths
        let w2DataToMap = extractedData.extractedData || extractedData;
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Data to map to 1040:`, JSON.stringify(w2DataToMap, null, 2));
          console.log(`🔍 [1040 GET] Current form1040Data before mapping:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check for personal info fields in the W2 data
        console.log(`🔍 [1040 GET DEBUG] Checking W2 personal info fields:`);
        console.log(`  - employeeName: ${w2DataToMap.employeeName}`);
        console.log(`  - Employee?.Name: ${w2DataToMap.Employee?.Name}`);
        console.log(`  - Employee.Name: ${w2DataToMap['Employee.Name']}`);
        console.log(`  - employeeSSN: ${w2DataToMap.employeeSSN}`);
        console.log(`  - Employee?.SSN: ${w2DataToMap.Employee?.SSN}`);
        console.log(`  - Employee.SSN: ${w2DataToMap['Employee.SSN']}`);
        console.log(`  - employeeAddress: ${w2DataToMap.employeeAddress}`);
        console.log(`  - Employee?.Address: ${w2DataToMap.Employee?.Address}`);
        console.log(`  - Employee.Address: ${w2DataToMap['Employee.Address']}`);
        
        // Map W2 data to 1040 form fields
        const mappedData = W2ToForm1040Mapper.mapW2ToForm1040(
          w2DataToMap, 
          form1040Data
        );
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Mapped data from W2:`, JSON.stringify(mappedData, null, 2));
        }
        
        // DEBUG: Check if personalInfo was created in mappedData
        if (mappedData.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo was created:`, JSON.stringify(mappedData.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo was NOT created in mappedData`);
        }
        
        // Merge the mapped data
        form1040Data = { ...form1040Data, ...mappedData };
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Form1040Data after merging:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check if personalInfo exists in final form1040Data
        if (form1040Data.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo exists in final form1040Data:`, JSON.stringify(form1040Data.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo is MISSING from final form1040Data`);
        }
        
        // Create mapping summary
        const mappingSummary = W2ToForm1040Mapper.createMappingSummary(w2DataToMap);
        
        w2MappingData.push({
          documentId: w2Doc.id,
          fileName: w2Doc.fileName,
          mappings: mappingSummary
        });
        
        console.log(`✅ [1040 GET] Successfully processed W2 document: ${w2Doc.fileName}`);
      } else {
        console.log(`⚠️ [1040 GET] W2 document ${w2Doc.fileName} has no extractedData or invalid format`);
      }
    }

    // Fill in basic info from tax return if not already populated
    // IMPORTANT: Only fill from taxReturn if we don't have W2 personal info
    const hasW2PersonalInfo = form1040Data.personalInfo && (
      form1040Data.personalInfo.firstName || 
      form1040Data.personalInfo.lastName || 
      form1040Data.personalInfo.ssn || 
      form1040Data.personalInfo.address
    );
    
    console.log(`🔍 [1040 GET DEBUG] hasW2PersonalInfo: ${hasW2PersonalInfo}`);
    console.log(`🔍 [1040 GET DEBUG] form1040Data.firstName: ${form1040Data.firstName}`);
    
    if (!form1040Data.firstName && !hasW2PersonalInfo) {
      console.log("🔍 [1040 GET] No W2 personal info found, filling from taxReturn data");
      form1040Data.firstName = taxReturn.firstName || '';
      form1040Data.lastName = taxReturn.lastName || '';
      form1040Data.ssn = taxReturn.ssn || '';
      form1040Data.spouseFirstName = taxReturn.spouseFirstName || undefined;
      form1040Data.spouseLastName = taxReturn.spouseLastName || undefined;
      form1040Data.spouseSSN = taxReturn.spouseSsn || undefined;
      form1040Data.address = taxReturn.address || '';
      form1040Data.city = taxReturn.city || '';
      form1040Data.state = taxReturn.state || '';
      form1040Data.zipCode = taxReturn.zipCode || '';
      form1040Data.filingStatus = taxReturn.filingStatus as any;
      form1040Data.taxYear = taxReturn.taxYear;
    } else if (hasW2PersonalInfo) {
      console.log("✅ [1040 GET] W2 personal info exists, preserving it and ensuring top-level fields are set");
      // Ensure top-level fields are set from W2 data if they exist
      if (!form1040Data.firstName && form1040Data.personalInfo?.firstName) {
        form1040Data.firstName = form1040Data.personalInfo.firstName;
      }
      if (!form1040Data.lastName && form1040Data.personalInfo?.lastName) {
        form1040Data.lastName = form1040Data.personalInfo.lastName;
      }
      if (!form1040Data.ssn && form1040Data.personalInfo?.ssn) {
        form1040Data.ssn = form1040Data.personalInfo.ssn;
      }
      if (!form1040Data.address && form1040Data.personalInfo?.address) {
        form1040Data.address = form1040Data.personalInfo.address;
      }
      if (!form1040Data.city && form1040Data.personalInfo?.city) {
        form1040Data.city = form1040Data.personalInfo.city;
      }
      if (!form1040Data.state && form1040Data.personalInfo?.state) {
        form1040Data.state = form1040Data.personalInfo.state;
      }
      if (!form1040Data.zipCode && form1040Data.personalInfo?.zipCode) {
        form1040Data.zipCode = form1040Data.personalInfo.zipCode;
      }
      
      // Set other required fields from taxReturn
      form1040Data.filingStatus = form1040Data.filingStatus || taxReturn.filingStatus as any;
      form1040Data.taxYear = form1040Data.taxYear || taxReturn.taxYear;
    } else {
      console.log("🔍 [1040 GET] Top-level firstName exists, filling missing fields from taxReturn");
      // Fill in missing fields from taxReturn without overriding existing data
      form1040Data.lastName = form1040Data.lastName || taxReturn.lastName || '';
      form1040Data.ssn = form1040Data.ssn || taxReturn.ssn || '';
      form1040Data.spouseFirstName = form1040Data.spouseFirstName || taxReturn.spouseFirstName || undefined;
      form1040Data.spouseLastName = form1040Data.spouseLastName || taxReturn.spouseLastName || undefined;
      form1040Data.spouseSSN = form1040Data.spouseSSN || taxReturn.spouseSsn || undefined;
      form1040Data.address = form1040Data.address || taxReturn.address || '';
      form1040Data.city = form1040Data.city || taxReturn.city || '';
      form1040Data.state = form1040Data.state || taxReturn.state || '';
      form1040Data.zipCode = form1040Data.zipCode || taxReturn.zipCode || '';
      form1040Data.filingStatus = form1040Data.filingStatus || taxReturn.filingStatus as any;
      form1040Data.taxYear = form1040Data.taxYear || taxReturn.taxYear;
    }

    console.log("✅ [1040 GET] Successfully retrieved 1040 form data");
    if (process.env.NODE_ENV === 'development') {
      console.log("🔍 [1040 GET] Final form1040Data being returned:", JSON.stringify(form1040Data, null, 2));
      console.log("🔍 [1040 GET] W2 mapping data being returned:", JSON.stringify(w2MappingData, null, 2));
    }
    
    // FINAL DEBUG: Ensure personalInfo is preserved in the response
    if (form1040Data.personalInfo) {
      console.log("✅ [1040 GET FINAL] personalInfo is present in final response:", JSON.stringify(form1040Data.personalInfo, null, 2));
    } else {
      console.log("❌ [1040 GET FINAL] personalInfo is MISSING from final response!");
      console.log("🔍 [1040 GET FINAL] Available keys in form1040Data:", Object.keys(form1040Data));
    }
    
    return NextResponse.json({
      form1040Data,
      w2MappingData,
      taxReturn: {
        id: taxReturn.id,
        taxYear: taxReturn.taxYear,
        filingStatus: taxReturn.filingStatus
      }
    });

  } catch (error) {
    console.error("💥 [1040 GET] Error retrieving form 1040 data:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}

// POST: Save 1040 form data
export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 POST] Starting form 1040 data save for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    const body = await request.json();
    const { form1040Data }: { form1040Data: Form1040Data } = body;

    // Verify tax return ownership
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      }
    });

    if (!taxReturn) {
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    // Update tax return with 1040 form data
    const updatedTaxReturn = await prisma.taxReturn.update({
      where: { id: params.id },
      data: {
        firstName: form1040Data.firstName,
        lastName: form1040Data.lastName,
        ssn: form1040Data.ssn,
        spouseFirstName: form1040Data.spouseFirstName,
        spouseLastName: form1040Data.spouseLastName,
        spouseSsn: form1040Data.spouseSSN,
        address: form1040Data.address,
        city: form1040Data.city,
        state: form1040Data.state,
        zipCode: form1040Data.zipCode,
        filingStatus: form1040Data.filingStatus as any,
        
        // Tax calculation fields
        totalIncome: form1040Data.line9,
        adjustedGrossIncome: form1040Data.line11,
        standardDeduction: form1040Data.line12,
        taxableIncome: form1040Data.line15,
        taxLiability: form1040Data.line16,
        totalWithholdings: form1040Data.line25a,
        refundAmount: form1040Data.line33,
        amountOwed: form1040Data.line37,
        
        lastSavedAt: new Date()
      }
    });

    // Store full 1040 form data as JSON in a custom field or separate table
    // For now, we'll store it as extractedData in a document record
    
    // First, try to find an existing 1040 document
    const existingForm1040Doc = await prisma.document.findFirst({
      where: {
        taxReturnId: params.id,
        documentType: 'OTHER_TAX_DOCUMENT',
        fileName: { contains: 'Form_1040' }
      }
    });

    let form1040Document;
    if (existingForm1040Doc) {
      // Update existing document
      form1040Document = await prisma.document.update({
        where: { id: existingForm1040Doc.id },
        data: {
          extractedData: form1040Data as any,
          processingStatus: 'COMPLETED',
          fileName: `Form_1040_${form1040Data.taxYear}.json`
        }
      });
    } else {
      // Create new document
      form1040Document = await prisma.document.create({
        data: {
          taxReturnId: params.id,
          fileName: `Form_1040_${form1040Data.taxYear}.json`,
          fileType: 'application/json',
          fileSize: JSON.stringify(form1040Data).length,
          filePath: '',
          documentType: 'OTHER_TAX_DOCUMENT',
          processingStatus: 'COMPLETED',
          extractedData: form1040Data as any
        }
      });
    }

    console.log("✅ [1040 POST] Successfully saved 1040 form data");
    
    return NextResponse.json({
      success: true,
      taxReturn: updatedTaxReturn,
      form1040Document: form1040Document
    });

  } catch (error) {
    console.error("💥 [1040 POST] Error saving form 1040 data:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
