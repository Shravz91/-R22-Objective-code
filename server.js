const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Multer setup for file uploads
const upload = multer({ storage: multer.memoryStorage() });

// Helper function to shuffle an array (Fisher-Yates shuffle)
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

app.post('/api/generate', upload.single('excelFile'), async (req, res) => {
  try {
    const { paperType } = req.body;

    if (!req.file) {
      return res.status(400).json({ error: 'No Excel file uploaded' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    const questionKey = Object.keys(jsonData[0]).find(key => key.trim() === 'Question');
    const typeKey = Object.keys(jsonData[0]).find(key => key.trim() === 'Type');

    if (!questionKey) {
      return res.status(400).json({ error: 'No "Question" column found in the Excel file' });
    }
    if (!typeKey) {
      return res.status(400).json({ error: 'No "Type" column found in the Excel file' });
    }

    // Process questions using the 'Type' column
    const questionBank = jsonData.map(row => {
      let questionText = row[questionKey] ? row[questionKey].toString().trim() : '';
      questionText = questionText.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' '); // Normalize to single space

      let type;
      if (row[typeKey] && row[typeKey].toLowerCase() === 'm') {
        type = 'multiple-choice';
      } else if (row[typeKey] && (row[typeKey].toLowerCase() === 'f' || row[typeKey].toLowerCase() === 'o')) {
        type = 'fill-in-the-blank';
      } else {
        console.log(`Skipping row due to invalid type: ${row[typeKey]}`);
        return null; // Skip invalid types
      }

      let question = questionText;
      let optionA = null;
      let optionB = null;
      let optionC = null;
      let optionD = null;

      if (type === 'multiple-choice') {
        // More flexible regex to handle variations in spacing and separators
        const optionRegex = /([A-Da-d])[\.\)]\s*(.*?)(?=\s*[A-Da-d][\.\)]\s*|$)/g;
        let match;
        let options = [];
        while ((match = optionRegex.exec(questionText)) !== null) {
          options.push({ letter: match[1].toUpperCase(), text: match[2].trim() });
        }

        if (options.length >= 4) {
          const firstMatch = questionText.match(/([A-Da-d])[\.\)]\s*/);
          if (firstMatch) {
            question = questionText.substring(0, firstMatch.index).trim();
            optionA = options.find(o => o.letter === 'A')?.text || null;
            optionB = options.find(o => o.letter === 'B')?.text || null;
            optionC = options.find(o => o.letter === 'C')?.text || null;
            optionD = options.find(o => o.letter === 'D')?.text || null;
          } else {
            console.log(`Failed to find first option match for question: ${questionText}`);
            return null;
          }
        } else {
          console.log(`Insufficient options (${options.length}) for multiple-choice question: ${questionText}`);
          return null;
        }
      }

      const unit = parseInt(row['Unit']) || 0;
      if (isNaN(row['Unit'])) {
        console.log(`Invalid unit value for question: ${questionText}, Unit: ${row['Unit']}`);
      }

      return {
        subjectCode: row['Subject Code'] || '',
        subject: row['Subject'] || '',
        branch: row['Branch'] || '',
        regulation: row['Regulation'] || '',
        year: row['Year'] || 0,
        semester: row['Sem'] || 0,
        month: row['Month'] || '',
        unit: unit,
        question: question,
        imageUrl: row['Image Url'] || null,
        type: type,
        ...(type === 'multiple-choice' ? { optionA, optionB, optionC, optionD } : {})
      };
    }).filter(q => {
      if (!q || !q.subjectCode || !q.question || q.unit <= 0) {
        console.log(`Filtering out invalid question: ${q?.question}, Unit: ${q?.unit}, SubjectCode: ${q?.subjectCode}`);
        return false;
      }
      return true;
    });

    if (questionBank.length === 0) {
      return res.status(400).json({ error: 'No valid questions found in the Excel file' });
    }

    // Organize questions by unit and type
    const multipleChoiceByUnit = {
      1: questionBank.filter(q => q.unit === 1 && q.type === 'multiple-choice'),
      2: questionBank.filter(q => q.unit === 2 && q.type === 'multiple-choice'),
      3: questionBank.filter(q => q.unit === 3 && q.type === 'multiple-choice'),
      4: questionBank.filter(q => q.unit === 4 && q.type === 'multiple-choice'),
      5: questionBank.filter(q => q.unit === 5 && q.type === 'multiple-choice')
    };

    const fillInTheBlankByUnit = {
      1: questionBank.filter(q => q.unit === 1 && q.type === 'fill-in-the-blank'),
      2: questionBank.filter(q => q.unit === 2 && q.type === 'fill-in-the-blank'),
      3: questionBank.filter(q => q.unit === 3 && q.type === 'fill-in-the-blank'),
      4: questionBank.filter(q => q.unit === 4 && q.type === 'fill-in-the-blank'),
      5: questionBank.filter(q => q.unit === 5 && q.type === 'fill-in-the-blank')
    };

    console.log('Multiple Choice Questions by Unit:', {
      Unit1: multipleChoiceByUnit[1].length,
      Unit2: multipleChoiceByUnit[2].length,
      Unit3: multipleChoiceByUnit[3].length,
      Unit4: multipleChoiceByUnit[4].length,
      Unit5: multipleChoiceByUnit[5].length
    });

    console.log('Fill-in-the-Blank Questions by Unit:', {
      Unit1: fillInTheBlankByUnit[1].length,
      Unit2: fillInTheBlankByUnit[2].length,
      Unit3: fillInTheBlankByUnit[3].length,
      Unit4: fillInTheBlankByUnit[4].length,
      Unit5: fillInTheBlankByUnit[5].length
    });

    let selectedQuestions = [];

    if (paperType === 'mid1') {
      // Mid 1: 2 MCQs + 2 FIBs from Unit 1, 2 MCQs + 2 FIBs from Unit 2, 1 MCQ + 1 FIB from Unit 3
      if (multipleChoiceByUnit[1].length < 2 || multipleChoiceByUnit[2].length < 2 || multipleChoiceByUnit[3].length < 1) {
        return res.status(400).json({
          error: `Insufficient multiple-choice questions for Mid 1: Need 2 from Unit 1 (found ${multipleChoiceByUnit[1].length}), 2 from Unit 2 (found ${multipleChoiceByUnit[2].length}), 1 from Unit 3 (found ${multipleChoiceByUnit[3].length})`
        });
      }
      if (fillInTheBlankByUnit[1].length < 2 || fillInTheBlankByUnit[2].length < 2 || fillInTheBlankByUnit[3].length < 1) {
        return res.status(400).json({
          error: `Insufficient fill-in-the-blank questions for Mid 1: Need 2 from Unit 1 (found ${fillInTheBlankByUnit[1].length}), 2 from Unit 2 (found ${fillInTheBlankByUnit[2].length}), 1 from Unit 3 (found ${fillInTheBlankByUnit[3].length})`
        });
      }

      selectedQuestions = [
        ...shuffleArray([...multipleChoiceByUnit[1]]).slice(0, 2), // 2 MCQs from Unit 1
        ...shuffleArray([...multipleChoiceByUnit[2]]).slice(0, 2), // 2 MCQs from Unit 2
        ...shuffleArray([...multipleChoiceByUnit[3]]).slice(0, 1), // 1 MCQ from Unit 3
        ...shuffleArray([...fillInTheBlankByUnit[1]]).slice(0, 2), // 2 FIBs from Unit 1
        ...shuffleArray([...fillInTheBlankByUnit[2]]).slice(0, 2), // 2 FIBs from Unit 2
        ...shuffleArray([...fillInTheBlankByUnit[3]]).slice(0, 1)  // 1 FIB from Unit 3
      ];

      console.log('Mid 1 Selection Breakdown:', {
        'Q1-Q2 (MC, Unit 1)': selectedQuestions.slice(0, 2).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q3-Q4 (MC, Unit 2)': selectedQuestions.slice(2, 4).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q5 (MC, Unit 3)': selectedQuestions.slice(4, 5).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q6-Q7 (FIB, Unit 1)': selectedQuestions.slice(5, 7).map(q => ({ question: q.question, unit: q.unit, type: q.type })),
        'Q8-Q9 (FIB, Unit 2)': selectedQuestions.slice(7, 9).map(q => ({ question: q.question, unit: q.unit, type: q.type })),
        'Q10 (FIB, Unit 3)': selectedQuestions.slice(9, 10).map(q => ({ question: q.question, unit: q.unit, type: q.type }))
      });
    } else if (paperType === 'mid2') {
      // Mid 2: 1 MCQ + 1 FIB from Unit 3, 2 MCQs + 2 FIBs from Unit 4, 2 MCQs + 2 FIBs from Unit 5
      if (multipleChoiceByUnit[3].length < 1 || multipleChoiceByUnit[4].length < 2 || multipleChoiceByUnit[5].length < 2) {
        return res.status(400).json({
          error: `Insufficient multiple-choice questions for Mid 2: Need 1 from Unit 3 (found ${multipleChoiceByUnit[3].length}), 2 from Unit 4 (found ${multipleChoiceByUnit[4].length}), 2 from Unit 5 (found ${multipleChoiceByUnit[5].length})`
        });
      }
      if (fillInTheBlankByUnit[3].length < 1 || fillInTheBlankByUnit[4].length < 2 || fillInTheBlankByUnit[5].length < 2) {
        return res.status(400).json({
          error: `Insufficient fill-in-the-blank questions for Mid 2: Need 1 from Unit 3 (found ${fillInTheBlankByUnit[3].length}), 2 from Unit 4 (found ${fillInTheBlankByUnit[4].length}), 2 from Unit 5 (found ${fillInTheBlankByUnit[5].length})`
        });
      }

      selectedQuestions = [
        ...shuffleArray([...multipleChoiceByUnit[3]]).slice(0, 1), // 1 MCQ from Unit 3
        ...shuffleArray([...multipleChoiceByUnit[4]]).slice(0, 2), // 2 MCQs from Unit 4
        ...shuffleArray([...multipleChoiceByUnit[5]]).slice(0, 2), // 2 MCQs from Unit 5
        ...shuffleArray([...fillInTheBlankByUnit[3]]).slice(0, 1), // 1 FIB from Unit 3
        ...shuffleArray([...fillInTheBlankByUnit[4]]).slice(0, 2), // 2 FIBs from Unit 4
        ...shuffleArray([...fillInTheBlankByUnit[5]]).slice(0, 2)  // 2 FIBs from Unit 5
      ];

      console.log('Mid 2 Selection Breakdown:', {
        'Q1 (MC, Unit 3)': selectedQuestions.slice(0, 1).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q2-Q3 (MC, Unit 4)': selectedQuestions.slice(1, 3).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q4-Q5 (MC, Unit 5)': selectedQuestions.slice(3, 5).map(q => ({ question: q.question, unit: q.unit, type: q.type, options: [q.optionA, q.optionB, q.optionC, q.optionD] })),
        'Q6 (FIB, Unit 3)': selectedQuestions.slice(5, 6).map(q => ({ question: q.question, unit: q.unit, type: q.type })),
        'Q7-Q8 (FIB, Unit 4)': selectedQuestions.slice(6, 8).map(q => ({ question: q.question, unit: q.unit, type: q.type })),
        'Q9-Q10 (FIB, Unit 5)': selectedQuestions.slice(8, 10).map(q => ({ question: q.question, unit: q.unit, type: q.type }))
      });
    } else {
      return res.status(400).json({ error: 'Invalid paperType. Use "mid1" or "mid2".' });
    }

    const paperDetails = {
      subjectCode: selectedQuestions[0].subjectCode,
      subject: selectedQuestions[0].subject,
      branch: selectedQuestions[0].branch,
      regulation: selectedQuestions[0].regulation,
      year: selectedQuestions[0].year,
      semester: selectedQuestions[0].semester,
      month: selectedQuestions[0].month
    };

    const response = {
      paperDetails,
      questions: selectedQuestions.map(q => ({
        question: q.question,
        unit: q.unit,
        imageUrl: q.imageUrl,
        type: q.type,
        ...(q.type === 'multiple-choice' ? { optionA: q.optionA, optionB: q.optionB, optionC: q.optionC, optionD: q.optionD } : {})
      }))
    };

    res.status(200).json(response);
  } catch (error) {
    console.error('Error generating questions:', error);
    res.status(500).json({ error: 'Error generating question paper: ' + error.message });
  }
});

app.get('/api/image-proxy-base64', async (req, res) => {
  const { url } = req.query;
  if (!url) return res.status(400).json({ error: 'No image URL provided' });

  try {
    const fetch = (await import('node-fetch')).default;
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to fetch image');
    const buffer = await response.buffer();
    const base64 = buffer.toString('base64');
    const mimeType = response.headers.get('content-type') || 'image/png';
    const dataUrl = `data:${mimeType};base64,${base64}`;
    res.json({ dataUrl });
  } catch (error) {
    console.error('Error fetching image:', error);
    res.status(500).json({ error: 'Failed to fetch image: ' + error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
