function rgbToAssColor(rgbColor) {
  if (typeof rgbColor === 'string') {
    rgbColor = rgbColor.replace('#', '');
    if (rgbColor.length === 6) {
      const r = parseInt(rgbColor.substring(0, 2), 16);
      const g = parseInt(rgbColor.substring(2, 4), 16);
      const b = parseInt(rgbColor.substring(4, 6), 16);
      return `&H00${b.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${r.toString(16).padStart(2, '0')}`;
    }
  }
  return '&H00FFFFFF';
}

function formatAssTime(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = Math.floor(seconds % 60);
  const centiseconds = Math.round((seconds - Math.floor(seconds)) * 100);
  return `${hours}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}.${centiseconds.toString().padStart(2, '0')}`;
}

function processText(text, replaceDict, allCaps, maxWordsPerLine) {
  let processed = text;
  
  for (const [oldWord, newWord] of Object.entries(replaceDict)) {
    const regex = new RegExp(oldWord.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    processed = processed.replace(regex, newWord);
  }
  
  if (allCaps) {
    processed = processed.toUpperCase();
  }
  
  if (maxWordsPerLine > 0) {
    const words = processed.split(/\s+/);
    const lines = [];
    for (let i = 0; i < words.length; i += maxWordsPerLine) {
      lines.push(words.slice(i, i + maxWordsPerLine).join(' '));
    }
    processed = lines.join('\\N');
  }
  
  return processed;
}

function splitLines(text, maxWordsPerLine) {
  if (maxWordsPerLine <= 0) return [text];
  const words = text.trim().split(/\s+/);
  const lines = [];
  for (let i = 0; i < words.length; i += maxWordsPerLine) {
    lines.push(words.slice(i, i + maxWordsPerLine).join(' '));
  }
  return lines;
}

function determineAlignmentCode(position, alignment, x, y, videoWidth, videoHeight) {
  const horizontalMap = { left: 1, center: 2, right: 3 };
  
  if (x !== null && x !== undefined && y !== null && y !== undefined) {
    const horizCode = horizontalMap[alignment] || 2;
    const anCode = 4 + (horizCode - 1);
    return { anCode, usePos: true, finalX: x, finalY: y };
  }
  
  let verticalBase, verticalCenter;
  if (position.includes('top')) {
    verticalBase = 7;
    verticalCenter = videoHeight / 6;
  } else if (position.includes('middle')) {
    verticalBase = 4;
    verticalCenter = videoHeight / 2;
  } else {
    verticalBase = 1;
    verticalCenter = (5 * videoHeight) / 6;
  }
  
  let leftBoundary, rightBoundary, centerLine;
  if (position.includes('left')) {
    leftBoundary = 0;
    rightBoundary = videoWidth / 3;
    centerLine = videoWidth / 6;
  } else if (position.includes('right')) {
    leftBoundary = (2 * videoWidth) / 3;
    rightBoundary = videoWidth;
    centerLine = (5 * videoWidth) / 6;
  } else {
    leftBoundary = videoWidth / 3;
    rightBoundary = (2 * videoWidth) / 3;
    centerLine = videoWidth / 2;
  }
  
  let finalX;
  if (alignment === 'left') {
    finalX = leftBoundary;
  } else if (alignment === 'right') {
    finalX = rightBoundary;
  } else {
    finalX = centerLine;
  }
  
  const finalY = verticalCenter;
  const horizCode = horizontalMap[alignment] || 2;
  const anCode = verticalBase + (horizCode - 1);
  
  return { anCode, usePos: true, finalX: Math.round(finalX), finalY: Math.round(finalY) };
}

function createStyleLine(styleOptions, videoResolution) {
  const fontFamily = styleOptions.font_family || 'Arial';
  const fontSize = styleOptions.font_size || Math.round(videoResolution[1] * 0.05);
  const lineColor = rgbToAssColor(styleOptions.line_color || '#FFFFFF');
  const secondaryColor = lineColor;
  const outlineColor = rgbToAssColor(styleOptions.outline_color || '#000000');
  const boxColor = rgbToAssColor(styleOptions.box_color || '#000000');
  
  const bold = styleOptions.bold ? '1' : '0';
  const italic = styleOptions.italic ? '1' : '0';
  const underline = styleOptions.underline ? '1' : '0';
  const strikeout = styleOptions.strikeout ? '1' : '0';
  
  const scaleX = styleOptions.scale_x || '100';
  const scaleY = styleOptions.scale_y || '100';
  const spacing = styleOptions.spacing || '0';
  const angle = styleOptions.angle || '0';
  const borderStyle = styleOptions.border_style || '1';
  const outlineWidth = styleOptions.outline_width || '2';
  const shadowOffset = styleOptions.shadow_offset || '0';
  
  const marginL = styleOptions.margin_l || '20';
  const marginR = styleOptions.margin_r || '20';
  const marginV = styleOptions.margin_v || '20';
  const alignment = styleOptions.alignment_code || '5';
  
  return `Style: Default,${fontFamily},${fontSize},${lineColor},${secondaryColor},${outlineColor},${boxColor},${bold},${italic},${underline},${strikeout},${scaleX},${scaleY},${spacing},${angle},${borderStyle},${outlineWidth},${shadowOffset},${alignment},${marginL},${marginR},${marginV},0`;
}

function handleClassic(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'middle_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const events = [];
  
  for (const segment of transcriptionResult.segments) {
    const text = segment.text.trim().replace(/\n/g, ' ');
    const lines = splitLines(text, maxWordsPerLine);
    const processedText = lines.map(line => processText(line, replaceDict, allCaps, 0)).join('\\N');
    
    const startTime = formatAssTime(segment.start);
    const endTime = formatAssTime(segment.end);
    const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
    
    events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}${processedText}`);
  }
  
  return events.join('\n');
}

function handleKaraoke(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'middle_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const events = [];
  
  for (const segment of transcriptionResult.segments) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    let dialogueText = '';
    
    if (maxWordsPerLine > 0) {
      const lines = [];
      for (let i = 0; i < words.length; i += maxWordsPerLine) {
        const lineWords = words.slice(i, i + maxWordsPerLine);
        const lineContent = lineWords.map(w => {
          const durationCs = Math.round((w.end - w.start) * 100);
          const wordText = processText(w.word, replaceDict, allCaps, 0);
          return `{\\k${durationCs}}${wordText} `;
        }).join('').trim();
        lines.push(lineContent);
      }
      dialogueText = lines.join('\\N');
    } else {
      dialogueText = words.map(w => {
        const durationCs = Math.round((w.end - w.start) * 100);
        const wordText = processText(w.word, replaceDict, allCaps, 0);
        return `{\\k${durationCs}}${wordText} `;
      }).join('').trim();
    }
    
    const startTime = formatAssTime(words[0].start);
    const endTime = formatAssTime(words[words.length - 1].end);
    const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
    
    events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}{\\c${wordColor}}${dialogueText}`);
  }
  
  return events.join('\n');
}

function handleHighlight(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'middle_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const lineColor = rgbToAssColor(styleOptions.line_color || '#FFFFFF');
  
  const events = [];
  
  for (const segment of transcriptionResult.segments) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    const processedWords = words.map(w => ({
      text: processText(w.word, replaceDict, allCaps, 0),
      start: w.start,
      end: w.end
    })).filter(w => w.text);
    
    if (!processedWords.length) continue;
    
    const lineSets = maxWordsPerLine > 0 
      ? processedWords.reduce((acc, word, idx) => {
          const lineIdx = Math.floor(idx / maxWordsPerLine);
          if (!acc[lineIdx]) acc[lineIdx] = [];
          acc[lineIdx].push(word);
          return acc;
        }, [])
      : [processedWords];
    
    for (const lineSet of lineSets) {
      const baseText = lineSet.map(w => w.text).join(' ');
      const lineStart = lineSet[0].start;
      const lineEnd = lineSet[lineSet.length - 1].end;
      
      const startTime = formatAssTime(lineStart);
      const endTime = formatAssTime(lineEnd);
      const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
      
      events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}{\\c${lineColor}}${baseText}`);
      
      for (let idx = 0; idx < lineSet.length; idx++) {
        const highlightedWords = lineSet.map((w, i) => 
          i === idx ? `{\\c${wordColor}}${w.text}{\\c${lineColor}}` : w.text
        ).join(' ');
        
        const wordStart = formatAssTime(lineSet[idx].start);
        const wordEnd = formatAssTime(lineSet[idx].end);
        
        events.push(`Dialogue: 1,${wordStart},${wordEnd},Default,,0,0,0,,${positionTag}{\\c${lineColor}}${highlightedWords}`);
      }
    }
  }
  
  return events.join('\n');
}

function handleUnderline(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'middle_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const lineColor = rgbToAssColor(styleOptions.line_color || '#FFFFFF');
  const events = [];
  
  for (const segment of transcriptionResult.segments) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    const processedWords = words.map(w => ({
      text: processText(w.word, replaceDict, allCaps, 0),
      start: w.start,
      end: w.end
    })).filter(w => w.text);
    
    if (!processedWords.length) continue;
    
    const lineSets = maxWordsPerLine > 0 
      ? processedWords.reduce((acc, word, idx) => {
          const lineIdx = Math.floor(idx / maxWordsPerLine);
          if (!acc[lineIdx]) acc[lineIdx] = [];
          acc[lineIdx].push(word);
          return acc;
        }, [])
      : [processedWords];
    
    for (const lineSet of lineSets) {
      for (let idx = 0; idx < lineSet.length; idx++) {
        const lineWords = lineSet.map((w, i) => 
          i === idx ? `{\\u1}${w.text}{\\u0}` : w.text
        ).join(' ');
        
        const startTime = formatAssTime(lineSet[idx].start);
        const endTime = formatAssTime(lineSet[idx].end);
        const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
        
        events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}{\\c${lineColor}}${lineWords}`);
      }
    }
  }
  
  return events.join('\n');
}

function handleWordByWord(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'middle_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const events = [];
  
  for (const segment of transcriptionResult.segments) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    const wordGroups = maxWordsPerLine > 0 
      ? words.reduce((acc, word, idx) => {
          const groupIdx = Math.floor(idx / maxWordsPerLine);
          if (!acc[groupIdx]) acc[groupIdx] = [];
          acc[groupIdx].push(word);
          return acc;
        }, [])
      : [words];
    
    for (const wordGroup of wordGroups) {
      for (const wInfo of wordGroup) {
        const word = processText(wInfo.word, replaceDict, allCaps, 0);
        if (!word) continue;
        
        const startTime = formatAssTime(wInfo.start);
        const endTime = formatAssTime(wInfo.end);
        const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
        
        events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}{\\c${wordColor}}${word}`);
      }
    }
  }
  
  return events.join('\n');
}

const styleHandlers = {
  classic: handleClassic,
  karaoke: handleKaraoke,
  highlight: handleHighlight,
  underline: handleUnderline,
  word_by_word: handleWordByWord
};

function generateAssContent(transcriptionResult, styleType, settings, replaceList, videoResolution) {
  const defaultSettings = {
    line_color: '#FFFFFF',
    word_color: '#FFFF00',
    box_color: '#000000',
    outline_color: '#000000',
    all_caps: false,
    max_words_per_line: 0,
    font_size: null,
    font_family: 'Arial',
    bold: false,
    italic: false,
    underline: false,
    strikeout: false,
    outline_width: 2,
    shadow_offset: 0,
    border_style: 1,
    x: null,
    y: null,
    position: 'middle_center',
    alignment: 'center'
  };
  
  const styleOptions = { ...defaultSettings, ...settings };
  
  if (!styleOptions.font_size) {
    styleOptions.font_size = Math.round(videoResolution[1] * 0.05);
  }
  
  const header = `[Script Info]
ScriptType: v4.00+
PlayResX: ${videoResolution[0]}
PlayResY: ${videoResolution[1]}
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
${createStyleLine(styleOptions, videoResolution)}

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
`;
  
  const replaceDict = {};
  if (Array.isArray(replaceList)) {
    for (const item of replaceList) {
      if (item.find && item.replace) {
        replaceDict[item.find] = item.replace;
      }
    }
  }
  
  const handler = styleHandlers[styleType] || handleClassic;
  const dialogueLines = handler(transcriptionResult, styleOptions, replaceDict, videoResolution);
  
  return header + dialogueLines + '\n';
}

// Fungsi utama untuk n8n
function createAssFromWhisper(videoMetadata, transcriptionResult, captionStyle, replaceList) {
  const videoResolution = [videoMetadata.width, videoMetadata.height];
  const styleType = (captionStyle.style || 'classic').toLowerCase();
  
  return generateAssContent(
    transcriptionResult,
    styleType,
    captionStyle,
    replaceList || [],
    videoResolution
  );
}

// Contoh penggunaan dalam n8n Code Node:
// const assContent = createAssFromWhisper(
//   $input.first().json.videoMetadata,
//   $input.first().json.transcriptionResult,
//   $input.first().json.captionStyle,
//   $input.first().json.replaceList
// );
// return [{ json: { assContent } }];

module.exports = { createAssFromWhisper };
