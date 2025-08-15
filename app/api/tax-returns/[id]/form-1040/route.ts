
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
    for (const w2Doc of w2Documents) {
      if (w2Doc.extractedData && typeof w2Doc.extractedData === 'object') {
        const extractedData = w2Doc.extractedData as any;
        
        // Map W2 data to 1040 form fields
        const mappedData = W2ToForm1040Mapper.mapW2ToForm1040(
          extractedData.extractedData || extractedData, 
          form1040Data
        );
        
        // Merge the mapped data
        form1040Data = { ...form1040Data, ...mappedData };
        
        // Create mapping summary
        const mappingSummary = W2ToForm1040Mapper.createMappingSummary(
          extractedData.extractedData || extractedData
        );
        
        w2MappingData.push({
          documentId: w2Doc.id,
          fileName: w2Doc.fileName,
          mappings: mappingSummary
        });
      }
    }

    // Fill in basic info from tax return if not already populated
    if (!form1040Data.firstName) {
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
    }

    console.log("✅ [1040 GET] Successfully retrieved 1040 form data");
    
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
