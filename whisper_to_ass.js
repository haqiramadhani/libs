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
  let processed = text || '';
  
  for (const [oldWord, newWord] of Object.entries(replaceDict)) {
    if (oldWord && newWord) {
      const regex = new RegExp(oldWord.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      processed = processed.replace(regex, newWord);
    }
  }
  
  if (allCaps) {
    processed = processed.toUpperCase();
  }
  
  if (maxWordsPerLine > 0) {
    const words = processed.split(/\s+/).filter(w => w);
    const lines = [];
    for (let i = 0; i < words.length; i += maxWordsPerLine) {
      lines.push(words.slice(i, i + maxWordsPerLine).join(' '));
    }
    processed = lines.join('\\N');
  }
  
  return processed;
}

function splitLines(text, maxWordsPerLine) {
  if (maxWordsPerLine <= 0) return [text || ''];
  const words = (text || '').trim().split(/\s+/).filter(w => w);
  const lines = [];
  for (let i = 0; i < words.length; i += maxWordsPerLine) {
    lines.push(words.slice(i, i + maxWordsPerLine).join(' '));
  }
  return lines.length ? lines : [''];
}

function determineAlignmentCode(position, alignment, x, y, videoWidth, videoHeight) {
  const horizontalMap = { left: 1, center: 2, right: 3 };
  
  if (x !== null && x !== undefined && y !== null && y !== undefined) {
    const horizCode = horizontalMap[alignment] || 2;
    const anCode = 4 + (horizCode - 1);
    return { anCode, usePos: true, finalX: x, finalY: y };
  }
  
  const positionLower = (position || 'bottom_center').toLowerCase();
  
  let verticalBase, verticalCenter;
  if (positionLower.includes('top')) {
    verticalBase = 7;
    verticalCenter = Math.round(videoHeight / 6);
  } else if (positionLower.includes('middle')) {
    verticalBase = 4;
    verticalCenter = Math.round(videoHeight / 2);
  } else {
    verticalBase = 1;
    verticalCenter = Math.round((5 * videoHeight) / 6);
  }
  
  let leftBoundary, rightBoundary, centerLine;
  if (positionLower.includes('left')) {
    leftBoundary = 0;
    rightBoundary = Math.round(videoWidth / 3);
    centerLine = Math.round(videoWidth / 6);
  } else if (positionLower.includes('right')) {
    leftBoundary = Math.round((2 * videoWidth) / 3);
    rightBoundary = videoWidth;
    centerLine = Math.round((5 * videoWidth) / 6);
  } else {
    leftBoundary = Math.round(videoWidth / 3);
    rightBoundary = Math.round((2 * videoWidth) / 3);
    centerLine = Math.round(videoWidth / 2);
  }
  
  let finalX;
  if (alignment === 'left') {
    finalX = leftBoundary;
  } else if (alignment === 'right') {
    finalX = rightBoundary;
  } else {
    finalX = centerLine;
  }
  
  const horizCode = horizontalMap[alignment] || 2;
  const anCode = verticalBase + (horizCode - 1);
  
  return { anCode, usePos: true, finalX, finalY: verticalCenter };
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
  const alignment = '5';
  
  return `Style: Default,${fontFamily},${fontSize},${lineColor},${secondaryColor},${outlineColor},${boxColor},${bold},${italic},${underline},${strikeout},${scaleX},${scaleY},${spacing},${angle},${borderStyle},${outlineWidth},${shadowOffset},${alignment},${marginL},${marginR},${marginV},0`;
}

// ============================================
// CRITICAL: Preprocess transcription data to avoid duplicate words
// ============================================
function restructureTranscriptionResult(transcriptionResult) {
  // If no root-level words array, return as-is
  if (!transcriptionResult.words || !Array.isArray(transcriptionResult.words)) {
    return transcriptionResult;
  }
  
  // If segments already have words, return as-is
  const hasSegmentWords = (transcriptionResult.segments || []).some(s => s.words && s.words.length > 0);
  if (hasSegmentWords) {
    return transcriptionResult;
  }
  
  // Create a working copy with empty words arrays for each segment
  const result = {
    ...transcriptionResult,
    segments: (transcriptionResult.segments || []).map(s => ({ ...s, words: [] }))
  };
  
  const words = [...result.words].sort((a, b) => (a.start || 0) - (b.start || 0));
  const segments = result.segments;
  
  // Assign each word to EXACTLY ONE segment based on start time
  for (const word of words) {
    const wordStart = word.start || 0;
    
    // Find the segment where this word starts
    const targetSegment = segments.find(seg => {
      const segStart = seg.start || 0;
      const segEnd = seg.end || 0;
      return wordStart >= segStart && wordStart < segEnd;
    });
    
    if (targetSegment) {
      targetSegment.words.push(word);
    }
  }
  
  // Remove the root-level words array to prevent double-processing
  delete result.words;
  
  return result;
}

function handleClassic(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'bottom_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const events = [];
  
  for (const segment of transcriptionResult.segments || []) {
    const text = (segment.text || '').trim().replace(/\n/g, ' ');
    if (!text) continue;
    
    const lines = splitLines(text, maxWordsPerLine);
    const processedText = lines.map(line => processText(line, replaceDict, allCaps, 0)).join('\\N');
    
    const startTime = formatAssTime(segment.start || 0);
    const endTime = formatAssTime(segment.end || 0);
    const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
    
    events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}${processedText}`);
  }
  
  return events.join('\n');
}

function handleKaraoke(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'bottom_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const events = [];
  
  for (const segment of transcriptionResult.segments || []) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    let dialogueText = '';
    
    if (maxWordsPerLine > 0) {
      const lines = [];
      for (let i = 0; i < words.length; i += maxWordsPerLine) {
        const lineWords = words.slice(i, i + maxWordsPerLine);
        const lineContent = lineWords.map(w => {
          const durationCs = Math.round(((w.end || 0) - (w.start || 0)) * 100);
          const wordText = processText(w.word || '', replaceDict, allCaps, 0);
          return `{\\k${durationCs}}${wordText} `;
        }).join('').trim();
        lines.push(lineContent);
      }
      dialogueText = lines.join('\\N');
    } else {
      dialogueText = words.map(w => {
        const durationCs = Math.round(((w.end || 0) - (w.start || 0)) * 100);
        const wordText = processText(w.word || '', replaceDict, allCaps, 0);
        return `{\\k${durationCs}}${wordText} `;
      }).join('').trim();
    }
    
    const startTime = formatAssTime(words[0].start || 0);
    const endTime = formatAssTime(words[words.length - 1].end || 0);
    const positionTag = `{\\an${anCode}\\pos(${finalX},${finalY})}`;
    
    events.push(`Dialogue: 0,${startTime},${endTime},Default,,0,0,0,,${positionTag}{\\c${wordColor}}${dialogueText}`);
  }
  
  return events.join('\n');
}

function handleHighlight(transcriptionResult, styleOptions, replaceDict, videoResolution) {
  const maxWordsPerLine = styleOptions.max_words_per_line || 0;
  const allCaps = styleOptions.all_caps || false;
  
  const position = styleOptions.position || 'bottom_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const lineColor = rgbToAssColor(styleOptions.line_color || '#FFFFFF');
  
  const events = [];
  
  for (const segment of transcriptionResult.segments || []) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    const processedWords = words.map(w => ({
      text: processText(w.word || '', replaceDict, allCaps, 0),
      start: w.start || 0,
      end: w.end || 0
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
  
  const position = styleOptions.position || 'bottom_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const lineColor = rgbToAssColor(styleOptions.line_color || '#FFFFFF');
  const events = [];
  
  for (const segment of transcriptionResult.segments || []) {
    const words = segment.words || [];
    if (!words.length) continue;
    
    const processedWords = words.map(w => ({
      text: processText(w.word || '', replaceDict, allCaps, 0),
      start: w.start || 0,
      end: w.end || 0
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
  
  const position = styleOptions.position || 'bottom_center';
  const alignment = styleOptions.alignment || 'center';
  const x = styleOptions.x;
  const y = styleOptions.y;
  
  const { anCode, finalX, finalY } = determineAlignmentCode(
    position, alignment, x, y, videoResolution[0], videoResolution[1]
  );
  
  const wordColor = rgbToAssColor(styleOptions.word_color || '#FFFF00');
  const events = [];
  
  for (const segment of transcriptionResult.segments || []) {
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
        const word = processText(wInfo.word || '', replaceDict, allCaps, 0);
        if (!word) continue;
        
        const startTime = formatAssTime(wInfo.start || 0);
        const endTime = formatAssTime(wInfo.end || 0);
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
    position: 'bottom_center',
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
      if (item && item.find && item.replace) {
        replaceDict[item.find] = item.replace;
      }
    }
  }
  
  const handler = styleHandlers[styleType] || handleClassic;
  const dialogueLines = handler(transcriptionResult, styleOptions, replaceDict, videoResolution);
  
  return header + dialogueLines + '\n';
}

// ============================================
// Fungsi utama untuk n8n dengan preprocessing
// ============================================
function createAssFromWhisper(videoMetadata, transcriptionResult, captionStyle, replaceList) {
  if (!videoMetadata || !transcriptionResult) {
    throw new Error('Missing required parameters: videoMetadata and transcriptionResult are required');
  }
  
  // CRITICAL: Restructure data to ensure each word appears in exactly one segment
  const restructuredResult = restructureTranscriptionResult(transcriptionResult);
  
  const videoResolution = [videoMetadata.width || 1920, videoMetadata.height || 1080];
  const styleType = (captionStyle && captionStyle.style || 'classic').toLowerCase();
  
  return generateAssContent(
    restructuredResult,
    styleType,
    captionStyle || {},
    replaceList || [],
    videoResolution
  );
}

module.exports = { createAssFromWhisper };
