// PPT 美化插件核心逻辑

// 配色方案
const COLOR_SCHEMES = {
  professional: {
    primary: '#1a365d',
    secondary: '#2d4a7c',
    accent: '#ed8936',
    text: '#2d3748',
    lightText: '#718096',
    background: '#ffffff'
  },
  modern: {
    primary: '#6366f1',
    secondary: '#8b5cf6',
    accent: '#f59e0b',
    text: '#1f2937',
    lightText: '#6b7280',
    background: '#ffffff'
  },
  elegant: {
    primary: '#0f172a',
    secondary: '#334155',
    accent: '#0ea5e9',
    text: '#1e293b',
    lightText: '#64748b',
    background: '#ffffff'
  }
};

// 字体配置
const FONT_CONFIG = {
  title: { name: '微软雅黑', size: 36 },
  subtitle: { name: '微软雅黑', size: 24 },
  heading: { name: '微软雅黑', size: 28 },
  body: { name: '微软雅黑', size: 18 },
  caption: { name: '微软雅黑', size: 14 }
};

let slideCount = 0;
let isProcessing = false;
let settings = {
  aiEnabled: false,
  apiKey: '',
  apiBase: '',
  model: 'claude-sonnet-4-20250514'
};

// Office 初始化
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log('Office.js 已加载');
    loadSettings();
    updateSlideCount();
  }
});

// 加载设置
function loadSettings() {
  try {
    const saved = localStorage.getItem('ppt-beautify-settings');
    if (saved) {
      settings = { ...settings, ...JSON.parse(saved) };
      document.getElementById('aiToggle').checked = settings.aiEnabled;
      document.getElementById('apiKey').value = settings.apiKey || '';
      document.getElementById('apiBase').value = settings.apiBase || '';
      document.getElementById('aiModel').value = settings.model || 'claude-sonnet-4-20250514';
      toggleAI();
    }
  } catch (e) {
    console.error('加载设置失败:', e);
  }
}

// 保存设置
function saveSettings() {
  settings.apiKey = document.getElementById('apiKey').value;
  settings.apiBase = document.getElementById('apiBase').value;
  settings.model = document.getElementById('aiModel').value;
  localStorage.setItem('ppt-beautify-settings', JSON.stringify(settings));
}

// 切换 AI 开关
function toggleAI() {
  settings.aiEnabled = document.getElementById('aiToggle').checked;
  const aiSettings = document.getElementById('aiSettings');
  if (settings.aiEnabled) {
    aiSettings.classList.remove('hidden');
  } else {
    aiSettings.classList.add('hidden');
  }
  saveSettings();
}

// 更新幻灯片数量显示
async function updateSlideCount() {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      
      slideCount = slides.items.length;
      document.getElementById('slideCount').textContent = `当前 PPT 共 ${slideCount} 页`;
    });
  } catch (error) {
    console.error('获取幻灯片数量失败:', error);
  }
}

// 显示状态
function showStatus(message, type = 'info') {
  const status = document.getElementById('status');
  status.textContent = message;
  status.className = `status ${type}`;
  status.classList.remove('hidden');
}

// 更新进度条
function updateProgress(percent) {
  const container = document.getElementById('progressContainer');
  const bar = document.getElementById('progressBar');
  container.classList.remove('hidden');
  bar.style.width = `${percent}%`;
}

// 隐藏进度条
function hideProgress() {
  document.getElementById('progressContainer').classList.add('hidden');
}

// 提取 PPT 内容
async function extractPPTContent() {
  let content = [];
  
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();
      
      let slideContent = { index: i + 1, texts: [] };
      
      for (const shape of slide.shapes.items) {
        try {
          shape.load('type');
          await context.sync();
          
          if (shape.type === 'GeometricShape' || shape.type === 'TextBox') {
            const textFrame = shape.textFrame;
            textFrame.load('textRange, hasText');
            await context.sync();
            
            if (textFrame.hasText) {
              const textRange = textFrame.textRange;
              textRange.load('text');
              await context.sync();
              
              if (textRange.text && textRange.text.trim()) {
                slideContent.texts.push(textRange.text.trim());
              }
            }
          }
        } catch (e) {
          // 忽略
        }
      }
      
      content.push(slideContent);
    }
  });
  
  return content;
}

// 调用 AI 获取美化建议
async function getAIBeautifyInstructions(content) {
  const apiKey = settings.apiKey;
  const apiBase = settings.apiBase || 'https://api.anthropic.com';
  const model = settings.model;
  
  if (!apiKey) {
    throw new Error('请先配置 API Key');
  }
  
  const prompt = `你是专业的 PPT 设计师。分析以下 PPT 内容，为每页生成美化指令。

PPT 内容：
${JSON.stringify(content, null, 2)}

请返回 JSON 格式的美化指令，结构如下：
{
  "slides": [
    {
      "index": 1,
      "colorScheme": "professional|modern|elegant",
      "elements": [
        {
          "text": "原文本内容",
          "type": "title|heading|body|caption",
          "fontSize": 36,
          "bold": true,
          "color": "#1a365d"
        }
      ]
    }
  ]
}

设计原则：
1. 首页标题用大字号(36-44pt)，加粗，深色
2. 小标题用中等字号(24-28pt)，加粗
3. 正文用标准字号(18-20pt)
4. 配色统一，主色调一致
5. 根据内容选择合适的配色方案

只返回 JSON，不要其他内容。`;

  const response = await fetch(`${apiBase}/v1/messages`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: model,
      max_tokens: 4096,
      messages: [{ role: 'user', content: prompt }]
    })
  });
  
  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || '调用 AI 失败');
  }
  
  const data = await response.json();
  const text = data.content[0].text;
  
  // 提取 JSON
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('AI 返回格式错误');
  }
  
  return JSON.parse(jsonMatch[0]);
}

// 开始美化
async function startBeautify() {
  if (isProcessing) return;
  
  const btn = document.getElementById('beautifyBtn');
  const undoBtn = document.getElementById('undoBtn');
  
  // 获取选项
  const options = {
    font: document.getElementById('optFont').checked,
    color: document.getElementById('optColor').checked,
    layout: document.getElementById('optLayout').checked,
    align: document.getElementById('optAlign').checked
  };
  
  isProcessing = true;
  btn.disabled = true;
  btn.textContent = '⏳ 美化中...';
  
  try {
    if (settings.aiEnabled) {
      // AI 模式
      showStatus('正在分析 PPT 内容...', 'processing');
      updateProgress(10);
      
      const content = await extractPPTContent();
      updateProgress(30);
      
      showStatus('AI 正在生成美化方案...', 'processing');
      const instructions = await getAIBeautifyInstructions(content);
      updateProgress(60);
      
      showStatus('正在应用美化...', 'processing');
      await applyAIInstructions(instructions);
      updateProgress(100);
      
    } else {
      // 规则模式
      showStatus('正在分析 PPT 结构...', 'processing');
      updateProgress(10);
      
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        
        const totalSlides = slides.items.length;
        
        for (let i = 0; i < totalSlides; i++) {
          const slide = slides.items[i];
          const progress = 10 + ((i + 1) / totalSlides) * 80;
          
          showStatus(`正在美化第 ${i + 1}/${totalSlides} 页...`, 'processing');
          updateProgress(progress);
          
          slide.shapes.load('items');
          await context.sync();
          
          for (const shape of slide.shapes.items) {
            await beautifyShape(context, shape, options, i === 0);
          }
          
          await context.sync();
        }
      });
      
      updateProgress(100);
    }
    
    showStatus('✅ 美化完成！', 'success');
    undoBtn.classList.remove('hidden');
    
  } catch (error) {
    console.error('美化失败:', error);
    showStatus(`❌ ${error.message}`, 'error');
  } finally {
    isProcessing = false;
    btn.disabled = false;
    btn.textContent = '🎨 开始美化';
    setTimeout(hideProgress, 2000);
  }
}

// 应用 AI 指令
async function applyAIInstructions(instructions) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    
    for (const slideInstr of instructions.slides) {
      const slideIndex = slideInstr.index - 1;
      if (slideIndex >= slides.items.length) continue;
      
      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();
      
      for (const shape of slide.shapes.items) {
        try {
          shape.load('type');
          await context.sync();
          
          if (shape.type === 'GeometricShape' || shape.type === 'TextBox') {
            const textFrame = shape.textFrame;
            textFrame.load('textRange, hasText');
            await context.sync();
            
            if (textFrame.hasText) {
              const textRange = textFrame.textRange;
              textRange.load('text');
              await context.sync();
              
              const text = textRange.text?.trim();
              if (!text) continue;
              
              // 找到匹配的指令
              const elemInstr = slideInstr.elements?.find(e => 
                e.text && text.includes(e.text.substring(0, 20))
              );
              
              if (elemInstr) {
                if (elemInstr.fontSize) textRange.font.size = elemInstr.fontSize;
                if (elemInstr.bold !== undefined) textRange.font.bold = elemInstr.bold;
                if (elemInstr.color) textRange.font.color = elemInstr.color;
                textRange.font.name = '微软雅黑';
              }
            }
          }
        } catch (e) {
          console.warn('处理形状时出错:', e);
        }
      }
      
      await context.sync();
    }
  });
}

// 美化单个形状（规则模式）
async function beautifyShape(context, shape, options, isFirstSlide) {
  try {
    shape.load('type, textFrame');
    await context.sync();
    
    if (shape.type === 'GeometricShape' || shape.type === 'TextBox') {
      const textFrame = shape.textFrame;
      textFrame.load('textRange, hasText');
      await context.sync();
      
      if (textFrame.hasText) {
        const textRange = textFrame.textRange;
        textRange.load('text, font');
        await context.sync();
        
        const text = textRange.text || '';
        const textLength = text.length;
        
        if (options.font) {
          if (isFirstSlide && textLength < 50) {
            textRange.font.name = FONT_CONFIG.title.name;
            textRange.font.size = FONT_CONFIG.title.size;
            textRange.font.bold = true;
          } else if (textLength < 30) {
            textRange.font.name = FONT_CONFIG.heading.name;
            textRange.font.size = FONT_CONFIG.heading.size;
            textRange.font.bold = true;
          } else {
            textRange.font.name = FONT_CONFIG.body.name;
            textRange.font.size = FONT_CONFIG.body.size;
            textRange.font.bold = false;
          }
        }
        
        if (options.color) {
          const scheme = COLOR_SCHEMES.professional;
          if (isFirstSlide || textLength < 30) {
            textRange.font.color = scheme.primary;
          } else {
            textRange.font.color = scheme.text;
          }
        }
      }
    }
  } catch (e) {
    console.warn('处理形状时出错:', e);
  }
}

// 撤销更改
function undoChanges() {
  showStatus('请按 Ctrl+Z (Mac: Cmd+Z) 撤销更改', 'info');
}
