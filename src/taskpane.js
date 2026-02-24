// PPT 美化插件核心逻辑

// 配色方案
const COLOR_SCHEMES = {
  professional: {
    primary: '#1a365d',    // 深蓝
    secondary: '#2d4a7c',  // 中蓝
    accent: '#ed8936',     // 橙色
    text: '#2d3748',       // 深灰
    lightText: '#718096',  // 浅灰
    background: '#ffffff'
  },
  modern: {
    primary: '#6366f1',    // 靛蓝
    secondary: '#8b5cf6',  // 紫色
    accent: '#f59e0b',     // 琥珀
    text: '#1f2937',
    lightText: '#6b7280',
    background: '#ffffff'
  },
  elegant: {
    primary: '#0f172a',    // 深黑蓝
    secondary: '#334155',
    accent: '#0ea5e9',     // 天蓝
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

// Office 初始化
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log('Office.js 已加载');
    updateSlideCount();
  }
});

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
        
        // 加载幻灯片的形状
        slide.shapes.load('items');
        await context.sync();
        
        // 美化每个形状
        for (const shape of slide.shapes.items) {
          await beautifyShape(context, shape, options, i === 0);
        }
        
        await context.sync();
      }
    });
    
    updateProgress(100);
    showStatus('✅ 美化完成！', 'success');
    undoBtn.classList.remove('hidden');
    
  } catch (error) {
    console.error('美化失败:', error);
    showStatus(`❌ 美化失败: ${error.message}`, 'error');
  } finally {
    isProcessing = false;
    btn.disabled = false;
    btn.textContent = '🎨 开始美化';
    setTimeout(hideProgress, 2000);
  }
}

// 美化单个形状
async function beautifyShape(context, shape, options, isFirstSlide) {
  try {
    shape.load('type, textFrame');
    await context.sync();
    
    // 只处理有文本的形状
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
        
        // 判断文本类型并应用样式
        if (options.font) {
          if (isFirstSlide && textLength < 50) {
            // 首页标题
            textRange.font.name = FONT_CONFIG.title.name;
            textRange.font.size = FONT_CONFIG.title.size;
            textRange.font.bold = true;
          } else if (textLength < 30) {
            // 小标题
            textRange.font.name = FONT_CONFIG.heading.name;
            textRange.font.size = FONT_CONFIG.heading.size;
            textRange.font.bold = true;
          } else {
            // 正文
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
    // 忽略单个形状的错误，继续处理其他形状
    console.warn('处理形状时出错:', e);
  }
}

// 撤销更改
function undoChanges() {
  // Office.js 没有直接的撤销 API，提示用户使用 Ctrl+Z
  showStatus('请按 Ctrl+Z (Mac: Cmd+Z) 撤销更改', 'info');
}
