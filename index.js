const express = require("express");
const path = require("path");

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static("public"));

app.post("/", (req, res) => {
  const { content } = req.body;
  try {
    res.json({ success: true, content: content });
  } catch (error) {
    console.error("Error:", error);
    res.status(500).json({ success: false, error: "Excel generation failed" });
  }
});

app.listen(port, () => console.log(`App listening on port ${port}`));