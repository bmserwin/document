const express = require("express")
const multer = require("multer")
const fs = require("fs")
const PizZip = require("pizzip")
const Docxtemplater = require("docxtemplater")
const JSZip = require("jszip")
const cors = require("cors")

const app = express()
app.use(express.json())
app.use(cors())
app.use(express.static("public"))

const upload = multer({ dest: "templates/" })

// Ensure directories exist
if (!fs.existsSync("output")) {
    fs.mkdirSync("output")
}
if (!fs.existsSync("templates")) {
    fs.mkdirSync("templates")
}

// Upload template
app.post("/upload", upload.single("file"), (req, res) => {
    // Rename uploaded file to template.docx so generator can find it
    fs.renameSync(req.file.path, "templates/template.docx")
    res.json({ message: "Template uploaded" })
})

// Generate documents
app.post("/generate", async (req, res) => {

    const { students } = req.body
    const templatePath = "templates/template.docx"

    const zipOutput = new JSZip()

    const content = fs.readFileSync(templatePath, "binary")

    for (let student of students) {

        const zip = new PizZip(content)
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true })

        doc.setData(student)

        try {
            doc.render()
        } catch (error) {
            console.error(error)
        }

        const buf = doc.getZip().generate({ type: "nodebuffer" })

        const filename = `${student.student_name}_${student.roll_number}.docx`

        zipOutput.file(filename, buf)
    }

    const zipFile = await zipOutput.generateAsync({ type: "nodebuffer" })

    fs.writeFileSync("output/docs.zip", zipFile)

    res.download("output/docs.zip")
})

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
