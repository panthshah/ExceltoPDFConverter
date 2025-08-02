"use client"

import type React from "react"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { FileSpreadsheet, FileDown, Eye } from "lucide-react"
import { read, utils } from "xlsx"
import { jsPDF } from "jspdf"
import autoTable from "jspdf-autotable"

export default function ExcelToPdfConverter() {
  const [file, setFile] = useState<File | null>(null)
  const [fileName, setFileName] = useState<string>("")
  const [pdfUrl, setPdfUrl] = useState<string | null>(null)
  const [isConverting, setIsConverting] = useState<boolean>(false)
  const [tableData, setTableData] = useState<any[] | null>(null)
  const [tableHeaders, setTableHeaders] = useState<string[] | null>(null)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0]
      setFile(selectedFile)
      setFileName(selectedFile.name)
      setPdfUrl(null)
      setTableData(null)
      setTableHeaders(null)
    }
  }

  const convertToPdf = async () => {
    if (!file) return

    setIsConverting(true)

    try {
      // Read the Excel file
      const data = await file.arrayBuffer()
      const workbook = read(data)

      // Get the first worksheet
      const worksheet = workbook.Sheets[workbook.SheetNames[0]]

      // Convert to JSON
      const jsonData = utils.sheet_to_json(worksheet)

      // Extract headers
      const headers = Object.keys(jsonData[0] || {})
      setTableHeaders(headers)
      setTableData(jsonData)

      // Create PDF
      const pdf = new jsPDF()

      // Add title
      pdf.setFontSize(16)
      pdf.text(fileName.replace(".xlsx", "").replace(".xls", ""), 14, 15)

      // Add table
      autoTable(pdf, {
        head: [headers],
        body: jsonData.map((row) => headers.map((header) => row[header])),
        startY: 25,
        theme: "grid",
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [41, 128, 185], textColor: 255 },
      })

      // Generate PDF URL
      const pdfBlob = pdf.output("blob")
      const url = URL.createObjectURL(pdfBlob)
      setPdfUrl(url)
    } catch (error) {
      console.error("Error converting file:", error)
    } finally {
      setIsConverting(false)
    }
  }

  return (
    <div className="container mx-auto py-10 px-4">
      <Card className="max-w-2xl mx-auto">
        <CardHeader>
          <CardTitle className="text-2xl">Excel to PDF Converter</CardTitle>
          <CardDescription>Upload an Excel file and convert it to a PDF with tabular data</CardDescription>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="space-y-2">
            <Label htmlFor="excel-file">Upload Excel File</Label>
            <div className="flex items-center gap-4">
              <div className="grid w-full max-w-sm items-center gap-1.5">
                <Input
                  id="excel-file"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  className="cursor-pointer"
                />
              </div>
              {file && (
                <Button onClick={convertToPdf} disabled={isConverting} className="whitespace-nowrap">
                  {isConverting ? "Converting..." : "Convert to PDF"}
                </Button>
              )}
            </div>
          </div>

          {file && (
            <div className="flex items-center gap-2 text-sm text-muted-foreground">
              <FileSpreadsheet className="h-4 w-4" />
              <span>{fileName}</span>
            </div>
          )}

          {tableData && tableHeaders && (
            <div className="border rounded-md overflow-auto max-h-[300px]">
              <table className="w-full border-collapse">
                <thead className="bg-muted sticky top-0">
                  <tr>
                    {tableHeaders.map((header, index) => (
                      <th key={index} className="p-2 text-left text-xs font-medium border">
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {tableData.slice(0, 10).map((row, rowIndex) => (
                    <tr key={rowIndex} className={rowIndex % 2 === 0 ? "bg-white" : "bg-muted/30"}>
                      {tableHeaders.map((header, colIndex) => (
                        <td key={colIndex} className="p-2 text-xs border">
                          {row[header]?.toString() || ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {tableData.length > 10 && (
                <div className="p-2 text-center text-xs text-muted-foreground">
                  Showing 10 of {tableData.length} rows
                </div>
              )}
            </div>
          )}
        </CardContent>

        {pdfUrl && (
          <CardFooter className="flex flex-col gap-4">
            <div className="flex gap-2 w-full">
              <Button variant="outline" className="flex-1" onClick={() => window.open(pdfUrl, "_blank")}>
                <Eye className="mr-2 h-4 w-4" />
                Preview PDF
              </Button>
              <Button
                className="flex-1"
                onClick={() => {
                  const link = document.createElement("a")
                  link.href = pdfUrl
                  link.download = fileName.replace(".xlsx", "").replace(".xls", "") + ".pdf"
                  link.click()
                }}
              >
                <FileDown className="mr-2 h-4 w-4" />
                Download PDF
              </Button>
            </div>
          </CardFooter>
        )}
      </Card>
    </div>
  )
}
