import Foundation
import PDFKit
import AppKit

let args = CommandLine.arguments
if args.count < 3 {
    print("Usage: render_pdf <input.pdf> <output.png> [page_index]")
    exit(1)
}

let pdfPath = args[1]
let outPath = args[2]
let pageIdx = args.count > 3 ? Int(args[3]) ?? 0 : 0

let pdfURL = URL(fileURLWithPath: pdfPath)
let outURL = URL(fileURLWithPath: outPath)

guard let doc = PDFDocument(url: pdfURL) else {
    print("Could not open PDF: \(pdfPath)")
    exit(1)
}

guard pageIdx < doc.pageCount, let page = doc.page(at: pageIdx) else {
    print("Invalid page index: \(pageIdx), max: \(doc.pageCount)")
    exit(1)
}

let pageRect = page.bounds(for: .mediaBox)
let scale: CGFloat = 2.0
let targetSize = NSSize(width: pageRect.width * scale, height: pageRect.height * scale)

let img = NSImage(size: targetSize)
img.lockFocus()
if let ctx = NSGraphicsContext.current?.cgContext {
    ctx.setFillColor(NSColor.white.cgColor)
    ctx.fill(CGRect(origin: .zero, size: targetSize))
    ctx.scaleBy(x: scale, y: scale)
    page.draw(with: .mediaBox, to: ctx)
}
img.unlockFocus()

if let tiff = img.tiffRepresentation,
   let rep = NSBitmapImageRep(data: tiff),
   let png = rep.representation(using: .png, properties: [:]) {
    try? png.write(to: outURL)
    print("Successfully rendered \(pdfPath) page \(pageIdx + 1) -> \(outPath)")
} else {
    print("Failed to encode PNG")
}
