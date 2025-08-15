
import { NextResponse } from "next/server"
import { getServerSession } from "next-auth"
import { authOptions } from "@/lib/auth"
import { prisma } from "@/lib/db"
import { getAzureDocumentIntelligenceService, type ExtractedFieldData } from "@/lib/azure-document-intelligence-service"
import { DuplicateDetectionService, type DuplicateDetectionResult } from "@/lib/duplicate-detection"

export const dynamic = "force-dynamic"

export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [PROCESS] Starting document processing for ID:", params.id)
  
  try {
    console.log("🔍 [PROCESS] Step 1: Getting server session...")
    const session = await getServerSession(authOptions)
    
    if (!session?.user?.email) {
      console.log("❌ [PROCESS] No session or email found")
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 })
    }
    console.log("✅ [PROCESS] Session found for email:", session.user.email)

    console.log("🔍 [PROCESS] Step 2: Finding user in database...")
    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    })

    if (!user) {
      console.log("❌ [PROCESS] User not found in database")
      return NextResponse.json({ error: "User not found" }, { status: 404 })
    }
    console.log("✅ [PROCESS] User found:", user.id)

    console.log("🔍 [PROCESS] Step 3: Finding document...")
    const document = await prisma.document.findFirst({
      where: { 
        id: params.id,
        taxReturn: {
          userId: user.id
        }
      }
    })

    if (!document) {
      console.log("❌ [PROCESS] Document not found or not owned by user")
      return NextResponse.json({ error: "Document not found" }, { status: 404 })
    }
    console.log("✅ [PROCESS] Document found:", {
      id: document.id,
      fileName: document.fileName,
      documentType: document.documentType,
      filePath: document.filePath
    })

    console.log("🔍 [PROCESS] Step 4: Updating document status to PROCESSING...")
    await prisma.document.update({
      where: { id: params.id },
      data: { processingStatus: 'PROCESSING' }
    })
    console.log("✅ [PROCESS] Document status updated to PROCESSING")

    console.log("🔍 [PROCESS] Step 5: Checking Azure Document Intelligence configuration...")
    const azureEndpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;
    const azureApiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY;
    
    if (!azureEndpoint || !azureApiKey) {
      console.log("❌ [PROCESS] Azure Document Intelligence not configured")
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: 'FAILED' }
      })
      return NextResponse.json(
        { error: "Document processing service not configured" },
        { status: 500 }
      )
    }
    
    console.log("✅ [PROCESS] Azure Document Intelligence configured")

    console.log("🔍 [PROCESS] Step 6: Processing with Azure Document Intelligence...")
    let extractedTaxData: any;
    
    try {
      const azureService = getAzureDocumentIntelligenceService();
      const extractedData = await azureService.extractDataFromDocument(document.filePath, document.documentType);
      
      extractedTaxData = {
        documentType: document.documentType,
        ocrText: extractedData.fullText || '',
        extractedData: extractedData,
        confidence: 0.95 // Azure typically has high confidence
      };
      console.log("✅ [PROCESS] Azure Document Intelligence processing completed")
    } catch (azureError) {
      console.error('❌ [PROCESS] Azure Document Intelligence processing failed:', azureError);
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: 'FAILED' }
      })
      return NextResponse.json(
        { error: "Document processing failed" },
        { status: 500 }
      )
    }

    console.log("🔍 [PROCESS] Step 7: Running duplicate detection...")
    let duplicateDetection: DuplicateDetectionResult;
    
    try {
      duplicateDetection = await DuplicateDetectionService.checkForDuplicates({
        documentType: extractedTaxData.documentType,
        extractedData: extractedTaxData.extractedData,
        taxReturnId: document.taxReturnId
      });
      
      console.log("✅ [PROCESS] Duplicate detection completed:", {
        isDuplicate: duplicateDetection.isDuplicate,
        confidence: duplicateDetection.confidence,
        matchingCount: duplicateDetection.matchingDocuments.length
      });
    } catch (duplicateError) {
      console.error("❌ [PROCESS] Duplicate detection failed:", duplicateError);
      // Continue processing even if duplicate detection fails
      duplicateDetection = {
        isDuplicate: false,
        confidence: 0,
        matchingDocuments: [],
        matchCriteria: {
          documentType: false,
          employerInfo: false,
          recipientInfo: false,
          amountSimilarity: false,
          nameSimilarity: false
        }
      };
    }

    console.log("🔍 [PROCESS] Step 8: Creating streaming response...")
    console.log("✅ [PROCESS] Extracted data preview:", {
      documentType: extractedTaxData.documentType,
      hasOcrText: !!extractedTaxData.ocrText,
      hasExtractedData: !!extractedTaxData.extractedData,
      confidence: extractedTaxData.confidence,
      duplicateFound: duplicateDetection.isDuplicate
    })

    // Create streaming response to maintain frontend compatibility
    const encoder = new TextEncoder()
    const readable = new ReadableStream({
      async start(controller) {
        try {
          console.log("🔍 [STREAM] Step 1: Preparing JSON response...")
          // Send the complete response as a single properly formatted chunk
          const jsonResponse = JSON.stringify({
            documentType: extractedTaxData.documentType,
            ocrText: extractedTaxData.ocrText,
            extractedData: extractedTaxData.extractedData,
            duplicateDetection: duplicateDetection
          });
          console.log("✅ [STREAM] JSON response prepared, length:", jsonResponse.length)

          console.log("🔍 [STREAM] Step 2: Sending content chunk...")
          // Send the complete JSON response as content
          controller.enqueue(encoder.encode(`data: ${JSON.stringify({content: jsonResponse})}\n\n`));
          console.log("✅ [STREAM] Content chunk sent")

          console.log("🔍 [STREAM] Step 3: Saving to database...")
          // Save OCR text and extracted data to database
          await prisma.document.update({
            where: { id: params.id },
            data: {
              ocrText: extractedTaxData.ocrText,
              extractedData: {
                documentType: extractedTaxData.documentType,
                ocrText: extractedTaxData.ocrText,
                extractedData: extractedTaxData.extractedData,
                confidence: extractedTaxData.confidence,
                duplicateDetection: JSON.parse(JSON.stringify(duplicateDetection)) // Convert to JSON-compatible format
              },
              processingStatus: 'COMPLETED'
            }
          })
          console.log("✅ [STREAM] Database update completed")

          console.log("🔍 [STREAM] Step 4: Sending completion signal...")
          controller.enqueue(encoder.encode('data: [DONE]\n\n'))
          controller.close()
          console.log("✅ [STREAM] Stream completed successfully")
        } catch (error) {
          console.error('💥 [STREAM] Error in streaming response:', error)
          console.error('💥 [STREAM] Error stack:', error instanceof Error ? error.stack : 'No stack trace')
          await prisma.document.update({
            where: { id: params.id },
            data: { processingStatus: 'FAILED' }
          })
          controller.error(error)
        }
      }
    })

    return new Response(readable, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-cache',
      },
    })

  } catch (error) {
    console.error("💥 [PROCESS] Document processing error:", error)
    console.error("💥 [PROCESS] Error stack:", error instanceof Error ? error.stack : 'No stack trace')
    
    // Update document status to failed
    try {
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: 'FAILED' }
      })
      console.log("✅ [PROCESS] Document status updated to FAILED")
    } catch (updateError) {
      console.error("💥 [PROCESS] Failed to update document status:", updateError)
    }

    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    )
  }
}


