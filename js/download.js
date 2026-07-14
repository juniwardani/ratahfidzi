
(function () {
  const btnDownload = document.getElementById('downloadBtn');
  if (!btnDownload) return;

  btnDownload.addEventListener('click', async () => {
    try {
      btnDownload.disabled = true;
      btnDownload.textContent = 'MENGUNDUH...';

      await downloadDocx();
    } catch (err) {
      console.error(err);
      alert('Gagal melakukan download DOCX: ' + (err?.message || err));
    } finally {
      btnDownload.disabled = false;
      btnDownload.textContent = '⬇️ Download';
    }
  });

  async function loadTemplate() {
    const res = await fetch('doc/template.docx');
    if (!res.ok) throw new Error('Gagal memuat template DOCX');
    const arrayBuffer = await res.arrayBuffer();
    return arrayBuffer;
  }

  function getSemester(tanggal) {
    if (!tanggal || typeof tanggal !== 'string') return '';
    const m = tanggal.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return '';
    const month = Number(m[2]);
    if (month >= 1 && month <= 6) return 'Genap';
    if (month >= 7 && month <= 12) return 'Ganjil';
    return '';
  }

  function formatTanggalIndonesia(tanggal) {
    if (!tanggal || typeof tanggal !== 'string') return '';
    const m = tanggal.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return tanggal;
    const year = Number(m[1]);
    const month = Number(m[2]);
    const day = Number(m[3]);
    const date = new Date(year, month - 1, day);
    const hari = [
      'Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'
    ];
    const bulan = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    return `${hari[date.getDay()]}, ${day} ${bulan[month - 1]} ${year}`;
  }

  function generateRows(penilaianArray, columnWidths) {
    let rowsXml = '';
    penilaianArray.forEach((p, i) => {
      rowsXml += `
        <w:tr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[0] || '1799'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${i + 1}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[1] || '1124'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${escapeXml(p.nama || '')}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[2] || '1382'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${escapeXml(p.kegiatan || '')}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[3] || '1394'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${escapeXml(p.indikator || '')}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[4] || '1405'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${escapeXml(p.penilaian || '')}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${columnWidths[5] || '2138'}" w:type="dxa"/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:r>
                <w:t>${escapeXml(p.catatan || '')}</w:t>
              </w:r>
            </w:p>
          </w:tc>
        </w:tr>
      `;
    });
    return rowsXml;
  }
  
  function extractColumnWidths(templateRow) {
    const widths = [];
    const tcStart = '<w:tc';
    const tcEnd = '</w:tc>';
    let pos = 0;
    
    while (pos < templateRow.length) {
      const tcStartIdx = templateRow.indexOf(tcStart, pos);
      if (tcStartIdx === -1) break;
      
      const tcEndIdx = templateRow.indexOf(tcEnd, tcStartIdx);
      if (tcEndIdx === -1) break;
      
      const tcContent = templateRow.slice(tcStartIdx, tcEndIdx);
      const widthMatch = tcContent.match(/<w:tcW w:w="(\d+)"/);
      if (widthMatch) {
        widths.push(widthMatch[1]);
      } else {
        widths.push('');
      }
      
      pos = tcEndIdx + tcEnd.length;
    }
    
    return widths;
  }

  async function createMultiPageDocx(dataArray) {
    const templateBuffer = await loadTemplate();
    const zip = new JSZip();
    await zip.loadAsync(templateBuffer);
    let documentXml = await zip.file('word/document.xml').async('string');

    // Extract the body template part
    const bodyStart = documentXml.indexOf('<w:body>') + '<w:body>'.length;
    const bodyEnd = documentXml.indexOf('<w:sectPr');
    const bodyTemplate = documentXml.slice(bodyStart, bodyEnd);

    // Generate content for each data entry
    let newBodyContent = '';
    dataArray.forEach((entry, index) => {
      let pageContent = bodyTemplate;

      // Replace placeholders
      // First, remove entire paragraphs that contain the {{#dates}}, {{/dates}}, {{#rows}}, {{/rows}}, {{^last}} placeholders
      pageContent = pageContent.replace(/<w:p[^>]*>[\s\S]*?\{\{#dates\}\}[\s\S]*?<\/w:p>/g, '');
      pageContent = pageContent.replace(/<w:p[^>]*>[\s\S]*?\{\{\/dates\}\}[\s\S]*?<\/w:p>/g, '');
      pageContent = pageContent.replace(/<w:p[^>]*>[\s\S]*?\{\{#rows\}\}[\s\S]*?<\/w:p>/g, '');
      pageContent = pageContent.replace(/<w:p[^>]*>[\s\S]*?\{\{\/rows\}\}[\s\S]*?<\/w:p>/g, '');
      pageContent = pageContent.replace(/<w:p[^>]*>[\s\S]*?\{\{\^last\}\}[\s\S]*?\{\{\/last\}\}[\s\S]*?<\/w:p>/g, '');
      
      // Then replace the actual content placeholders
      pageContent = pageContent.replace(/\{\{Topik\}\}/g, entry.topik || '');
      pageContent = pageContent.replace(/\{\{Subtopik\}\}/g, entry.subtopik || '');
      pageContent = pageContent.replace(/\{\{WaliKelas\}\}/g, entry.waliKelas || '');
      pageContent = pageContent.replace(/\{\{Kelas\}\}/g, entry.kelas || '');
      pageContent = pageContent.replace(/\{\{Tanggal\}\}/g, formatTanggalIndonesia(entry.tanggal) || '');
      pageContent = pageContent.replace(/\{\{Semester\}\}/g, getSemester(entry.tanggal) || '');

      // Replace rows part
      // First, extract column widths from the template row
      let columnWidths = [];
      const rowStartMarker = '<w:tr';
      const rowEndMarker = '</w:tr>';
      let rowStartIdx = pageContent.indexOf('<w:t>{{No}}</w:t>');
      
      if (rowStartIdx !== -1) {
        // Find the opening <w:tr> tag before {{No}}
        const actualRowStart = pageContent.slice(0, rowStartIdx).lastIndexOf(rowStartMarker);
        if (actualRowStart !== -1) {
          // Find the closing </w:tr> tag to get the full template row
          let depth = 1;
          let pos = actualRowStart + rowStartMarker.length;
          while (depth > 0 && pos < pageContent.length) {
            const nextTrStart = pageContent.indexOf('<w:tr', pos);
            const nextTrEnd = pageContent.indexOf(rowEndMarker, pos);
            
            if (nextTrStart !== -1 && nextTrStart < nextTrEnd) {
              depth++;
              pos = nextTrStart + 5;
            } else if (nextTrEnd !== -1) {
              depth--;
              pos = nextTrEnd + rowEndMarker.length;
            } else {
              break;
            }
          }
          const actualRowEnd = pos;
          const templateRow = pageContent.slice(actualRowStart, actualRowEnd);
          
          // Extract column widths
          columnWidths = extractColumnWidths(templateRow);
        }
      }
      
      // Generate new rows with the extracted column widths
      const rowsXml = generateRows(entry.penilaian || [], columnWidths);
      
      // Now replace the template row with our rows
      rowStartIdx = pageContent.indexOf('<w:t>{{No}}</w:t>');
      while (rowStartIdx !== -1) {
        const actualRowStart = pageContent.slice(0, rowStartIdx).lastIndexOf(rowStartMarker);
        if (actualRowStart === -1) break;
        
        let depth = 1;
        let pos = actualRowStart + rowStartMarker.length;
        while (depth > 0 && pos < pageContent.length) {
          const nextTrStart = pageContent.indexOf('<w:tr', pos);
          const nextTrEnd = pageContent.indexOf(rowEndMarker, pos);
          
          if (nextTrStart !== -1 && nextTrStart < nextTrEnd) {
            depth++;
            pos = nextTrStart + 5;
          } else if (nextTrEnd !== -1) {
            depth--;
            pos = nextTrEnd + rowEndMarker.length;
          } else {
            break;
          }
        }
        const actualRowEnd = pos;
        pageContent = pageContent.slice(0, actualRowStart) + rowsXml + pageContent.slice(actualRowEnd);
        rowStartIdx = -1;
      }

      // Remove any empty paragraphs left over
      pageContent = pageContent.replace(/<w:p\s*\/>/g, '');
      pageContent = pageContent.replace(/<w:p[^>]*>\s*<\/w:p>/g, '');

      newBodyContent += pageContent;

      // Add page break except for last entry
      if (index < dataArray.length - 1) {
        newBodyContent += `
          <w:p>
            <w:r>
              <w:br w:type="page"/>
            </w:r>
          </w:p>
        `;
      }
    });

    // Replace the body in documentXml
    documentXml = documentXml.slice(0, bodyStart) + newBodyContent + documentXml.slice(bodyEnd);

    // Save back to zip
    zip.file('word/document.xml', documentXml);

    // Generate blob
    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    return blob;
  }

  async function downloadDocx() {
    if (!window.currentUser) {
      throw new Error('Belum login.');
    }

    // Get filtered data
    const kelasData = window.data.filter(d => d.kelas === window.currentUser.kelas);
    let entries = kelasData;

    const filterTopik = document.getElementById('filterTopik')?.value || '';
    const filterSubtopik = document.getElementById('filterSubtopik')?.value || '';
    const filterTanggal = document.getElementById('filterTanggal')?.value || '';

    if (filterTopik) {
      entries = entries.filter(d => d.topik === filterTopik);
    }
    if (filterSubtopik) {
      entries = entries.filter(d => d.subtopik === filterSubtopik);
    }
    if (filterTanggal) {
      entries = entries.filter(d => d.tanggal === filterTanggal);
    }

    if (!entries || !entries.length) {
      throw new Error('Tidak ada data penilaian untuk di-download.');
    }

    const blob = await createMultiPageDocx(entries);

    // Download using FileSaver
    const filenameSafeKelas = (window.currentUser.kelas || 'kelas').replace(/[^a-z0-9\-_]+/gi, '_');
    saveAs(blob, `Rekap_${filenameSafeKelas}.docx`);
  }

  function escapeXml(unsafe) {
    return unsafe
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
})();
