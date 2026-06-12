import { useState, useRef } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
// @ts-ignore
import { saveAs } from 'file-saver';

// ========== 模板配置：在这里添加/修改模板 ==========
const TEMPLATE_OPTIONS = [
  { id: 'dot4',  label: '有封面型式检验报告',  file: './template_dot4.docx' },
  { id: 'dot3',  label: '无封面型式检验报告',  file: './template_dot3.docx' },
];
// ===================================================

export default function App() {
  const [loading, setLoading] = useState(false);
  const [selectedTemplate, setSelectedTemplate] = useState(TEMPLATE_OPTIONS[0].id);
  const [manualImages, setManualImages] = useState<{name: string, buffer: ArrayBuffer}[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleClearImages = () => {
    setManualImages([]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const fieldMapping: Record<string, string> = {
    "检验日期": "检验日期",
    "报告编号": "报告编号",
    "规格型号": "规格型号",
    "订单号": "订单号",
    "商标": "商标",
    "包装规格": "包装规格",
    "制动液型号": "制动液型号",
    "制动液名称": "制动液名称",
    "检验批次": "检验批次",
    "平衡回流沸点标准值": "平衡回流沸点标准值",
    "平衡回流沸点": "平衡回流沸点",
    "湿平衡回流沸点标准值": "湿平衡回流沸点标准值",
    "湿平衡回流沸点": "湿平衡回流沸点",
    "-40运动粘度标准值": "-40运动粘度标准值",
    "100运动粘度标准值": "100运动粘度标准值",
    "电导率": "电导率",
    "运动粘度-40": "运动粘度-40",
    "运动粘度100": "运动粘度100",
    "pH": "pH",
    "水含量": "水含量",
    "储备碱度": "储备碱度",
    "锌": "锌",
    "紫铜": "紫铜",
    "黄铜": "黄铜",
    "铸铁": "铸铁",
    "铝": "铝",
    "钢": "钢",
    "镀锡铁皮": "镀锡铁皮",
    "腐蚀性 pH": "腐蚀性 pH",
    "腐蚀性 试液沉淀物": "腐蚀性 试液沉淀物",
    "腐蚀性 硬度降低值": "腐蚀性 硬度降低值",
    "腐蚀性 根茎增值": "腐蚀性 根茎增值",
    "腐蚀性 体积增加值": "腐蚀性 体积增加值",
    "蒸发损失": "蒸发损失",
    "倾点": "倾点",
    "-40气泡上浮的时间": "-40气泡上浮的时间",
    "-50气泡上浮时间": "-50气泡上浮时间",
    "容水性气泡上浮时间": "容水性气泡上浮时间",
    "容水性沉淀量": "容水性沉淀量",
    "铝 抗氧化": "铝 抗氧化",
    "铸铁 抗氧化": "铸铁 抗氧化",
    "高温稳定性": "高温稳定性",
    "化学稳定性": "化学稳定性",
    "橡胶根茎增值": "橡胶根茎增值",
    "橡胶硬度降低值": "橡胶硬度降低值",
    "橡胶体积增加值": "橡胶体积增加值",
    "橡胶硬度降低值 试件": "橡胶硬度降低值 试件",
    "橡胶体积增加值 试件": "橡胶体积增加值 试件",
    "液体相容 沉淀体积": "液体相容 沉淀体积"
  };

  const toPluginBuffer = (ab: ArrayBuffer): Uint8Array => {
    const uint8 = new Uint8Array(ab);
    (uint8 as any).toString = function (encoding?: string) {
      if (encoding === 'hex') {
        let hex = '';
        for (let i = 0; i < Math.min(uint8.byteLength, 4); i++) hex += uint8[i].toString(16).padStart(2, '0');
        return hex;
      }
      return '[object Uint8Array]';
    };
    return uint8;
  };

  const handleManualFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
    const loadedImages = [];
    for (const file of files) {
      const buf = await file.arrayBuffer();
      loadedImages.push({ name: file.name, buffer: buf });
    }
    setManualImages(loadedImages);
  };

  const handleExport = async () => {
    setLoading(true);
    let currentStep = '初始化';

    try {
      currentStep = '连接飞书表格';
      const table = await bitable.base.getActiveTable();
      let recordId: string | null = null;
      try {
        const sel = await bitable.base.getSelection();
        recordId = sel?.recordId ?? null;
      } catch (_) {}

      if (!recordId) {
        alert('请先用鼠标点击选中一行数据后再生成报告！');
        setLoading(false);
        return;
      }

      currentStep = '读取字段数据';
      const data: Record<string, any> = {};

      for (const [key, fieldName] of Object.entries(fieldMapping)) {
        try {
          const field = await table.getFieldByName(fieldName);
          const val = await field.getValue(recordId);

          if (key.includes('日期') && val != null && val !== '') {
            try {
              let raw = val;
              if (Array.isArray(raw) && raw.length > 0) raw = raw[0]?.text ?? raw[0]?.value ?? raw;
              const d = new Date(Number(raw));
              data[key] = isNaN(d.getTime()) ? String(raw) : `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
            } catch { data[key] = String(val); }
            continue;
          }

          if (val == null || val === '') { data[key] = '/'; continue; }

          let str: string;
          if (typeof val === 'number' || typeof val === 'string') str = String(val);
          else if (Array.isArray(val)) str = val.map((v: any) => v?.text ?? v?.value ?? v?.name ?? String(v)).join('');
          else if (typeof val === 'object') str = String(val?.text ?? val?.value ?? val?.name ?? JSON.stringify(val));
          else str = String(val);

          const num = parseFloat(str);
          if (!isNaN(num) && str.trim() !== '') {
            if (key === '水含量') str = num.toFixed(3);
            else if (['液体相容 沉淀体积', '容水性沉淀量', '腐蚀性 试液沉淀物'].includes(key)) str = num.toFixed(2);
          }
          data[key] = str.trim() === '' ? '/' : str;

        } catch (fieldErr) {
          data[key] = '/';
        }
      }

      // 获取当前选中的模板文件路径
      currentStep = '加载 Word 模板';
      const templateConfig = TEMPLATE_OPTIONS.find(t => t.id === selectedTemplate)!;
      const res = await fetch(templateConfig.file);
      if (!res.ok) throw new Error(`找不到模板文件：${templateConfig.file}，请确认已上传到插件目录`);
      const templateBuf = await res.arrayBuffer();

      currentStep = '向 Word 模板注入文字数据';
      const zip = new PizZip(templateBuf);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '/' });
      doc.render(data);

      currentStep = '获取实验照片';
      // 优先手动上传；没有则自动从附件字段拉取
      let imagesToUse: { name: string; buffer: ArrayBuffer }[] = [];

      if (manualImages.length > 0) {
        imagesToUse = manualImages;
        console.log(`使用手动上传的 ${imagesToUse.length} 张图片`);
      } else {
        try {
          // ⚠️ 把「实验照片」改成你飞书表格里附件字段的实际名称
          const ATTACHMENT_FIELD_NAME = '实验照片';
          const attachFieldBase = await table.getFieldByName(ATTACHMENT_FIELD_NAME);
          const attachField = await table.getField(attachFieldBase.id) as any;
          const attachVal = await attachField.getValue(recordId) as any[];

          if (attachVal && attachVal.length > 0) {
            const sorted = [...attachVal].sort((a: any, b: any) =>
              (a.name ?? '').localeCompare(b.name ?? '', undefined, { numeric: true })
            );
            const urls = await attachField.getAttachmentUrls(recordId);
            for (let i = 0; i < sorted.length; i++) {
              try {
                const url = Array.isArray(urls) ? urls[i] : urls;
                if (!url) continue;
                const imgRes = await fetch(url);
                if (!imgRes.ok) throw new Error(`HTTP ${imgRes.status}`);
                const buf = await imgRes.arrayBuffer();
                imagesToUse.push({ name: sorted[i].name ?? `image_${i}`, buffer: buf });
                console.log(`✅ 自动获取附件：${sorted[i].name}`);
              } catch (e) {
                console.warn(`⚠️ 附件获取失败：${sorted[i].name}`, e);
              }
            }
          }
        } catch (attachErr) {
          console.warn('自动获取附件图片失败，跳过：', attachErr);
        }
      }

      currentStep = '批量物理替换图片';
      const renderedZip = doc.getZip();

      if (imagesToUse.length > 0) {
        let mediaFiles = Object.keys(renderedZip.files).filter(k =>
          k.startsWith('word/media/') &&
          k !== 'word/media/' &&
          !k.endsWith('/')
        );
        mediaFiles.sort((a, b) => {
          const numA = parseInt(a.match(/image(\d+)/)?.[1] || '0');
          const numB = parseInt(b.match(/image(\d+)/)?.[1] || '0');
          return numA - numB;
        });
        for (let i = 0; i < Math.min(imagesToUse.length, mediaFiles.length); i++) {
          const imgData = toPluginBuffer(imagesToUse[i].buffer);
          renderedZip.file(mediaFiles[i], imgData);
        }
      }

      currentStep = '生成文件下载';
      const out = renderedZip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      saveAs(out, `型式检验报告_${data['制动液名称'] ?? '新建'}_${data['检验批次'] ?? ''}.docx`);

    } catch (err: any) {
      alert(`出错步骤：【${currentStep}】\n报错信息：${err?.message ?? String(err)}`);
    } finally {
      setLoading(false);
    }
  };

  const onSafeExportClick = (e: any) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.nativeEvent && e.nativeEvent.stopImmediatePropagation) e.nativeEvent.stopImmediatePropagation();
    setTimeout(() => handleExport(), 0);
  };

  const currentLabel = TEMPLATE_OPTIONS.find(t => t.id === selectedTemplate)?.label ?? '';

  return (
    <div style={{ padding: '20px', fontFamily: 'sans-serif', maxWidth: '420px', margin: '0 auto' }}>
      <h3 style={{ marginBottom: '18px', color: '#1f2329', fontSize: '15px', textAlign: 'center' }}>
        型式检验报告生成
      </h3>

      {/* 模板选择区 */}
      <div style={{ marginBottom: '16px', padding: '14px', backgroundColor: '#f7f8fa', borderRadius: '8px', border: '1px solid #e0e4ea' }}>
        <p style={{ margin: '0 0 10px 0', fontSize: '13px', fontWeight: 'bold', color: '#1f2329' }}>
          📄 选择报告模板
        </p>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
          {TEMPLATE_OPTIONS.map(t => (
            <label
              key={t.id}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                padding: '10px 14px',
                borderRadius: '6px',
                border: `1.5px solid ${selectedTemplate === t.id ? '#1456f0' : '#dde1e7'}`,
                backgroundColor: selectedTemplate === t.id ? '#eef2ff' : '#ffffff',
                cursor: 'pointer',
                fontSize: '13px',
                color: '#1f2329',
                transition: 'all 0.15s',
              }}
            >
              <input
                type="radio"
                name="template"
                value={t.id}
                checked={selectedTemplate === t.id}
                onChange={() => setSelectedTemplate(t.id)}
                style={{ accentColor: '#1456f0' }}
              />
              {t.label}
            </label>
          ))}
        </div>

        {/* 已选提示 */}
        <p style={{ margin: '10px 0 0 0', fontSize: '12px', color: '#1456f0' }}>
          ✅ 当前使用：<b>{currentLabel}</b>
        </p>
      </div>

      {/* 图片上传区 */}
      <div style={{ marginBottom: '16px', padding: '14px', backgroundColor: '#f0f4ff', borderRadius: '8px', border: '1px dashed #1456f0' }}>
        <p style={{ margin: '0 0 6px 0', fontSize: '13px', color: '#1f2329' }}>
          🖼️ <b>实验照片</b>
        </p>
        <p style={{ margin: '0 0 10px 0', fontSize: '12px', color: '#666' }}>
          不上传则自动从表格附件字段读取；手动上传优先级更高，请命名为 01.jpg, 02.jpg...
        </p>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept="image/png, image/jpeg, image/jpg"
          onChange={handleManualFileChange}
          style={{ fontSize: '13px', width: '100%' }}
        />
        {manualImages.length > 0 && (
          <div style={{ marginTop: '10px', padding: '8px', backgroundColor: '#e1eaff', borderRadius: '4px', fontSize: '12px', color: '#1456f0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
              <b>✅ 已读取 {manualImages.length} 张照片：</b>
              <button
                onClick={handleClearImages}
                style={{ fontSize: '11px', padding: '2px 8px', backgroundColor: '#fff', color: '#e53935', border: '1px solid #e53935', borderRadius: '4px', cursor: 'pointer' }}
              >
                🗑️ 清除
              </button>
            </div>
            <ol style={{ margin: '0', paddingLeft: '18px', maxHeight: '80px', overflowY: 'auto' }}>
              {manualImages.map((img, idx) => <li key={idx}>{img.name}</li>)}
            </ol>
          </div>
        )}
      </div>

      {/* 导出按钮 */}
      <button
        onClick={onSafeExportClick}
        disabled={loading}
        style={{
          width: '100%',
          padding: '14px',
          fontSize: '15px',
          fontWeight: 'bold',
          backgroundColor: loading ? '#a4bcf6' : '#1456f0',
          color: '#fff',
          border: 'none',
          borderRadius: '8px',
          cursor: loading ? 'not-allowed' : 'pointer',
        }}
      >
        {loading ? '报告生成中...' : '🚀 导出检验报告'}
      </button>
    </div>
  );
}
