const { PDFParse } = require('pdf-parse');

async function downloadAndExtract(accessToken, hit) {
    const fileName = hit.resource?.name || "Unknown";
    const resourceId = hit.resource?.id;
    const driveId = hit.resource?.parentReference?.driveId;

    if (!driveId || !resourceId) {
        return hit.summary ? `[Source: ${fileName}]\n${hit.summary}` : null;
    }

    const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${resourceId}/content`;
    const downloadRes = await fetch(downloadUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!downloadRes.ok) {
        return hit.summary ? `[Source: ${fileName}]\n${hit.summary}` : null;
    }

    const fileBuffer = Buffer.from(await downloadRes.arrayBuffer());

    if (fileName.toLowerCase().endsWith('.pdf')) {
        const parser = new PDFParse({ data: fileBuffer });
        const pdfData = await parser.getText();
        await parser.destroy();

        const cleanText = pdfData.text
            .substring(0, 40000)
            .replace(/--\s*\d+\s*of\s*\d+\s*--/g, '')
            .replace(/CT\/EH\/\d+\.\d+/g, '')
            .replace(/--/g, '')
            .replace(/\.{3,}/g, '.')
            .replace(/\s+/g, ' ')
            .trim();

        return `[Source: ${fileName}]\n${cleanText}`;
    }

    return `[Source: ${fileName}]\n${fileBuffer.toString('utf8').substring(0, 15000)}`;
}

module.exports = { downloadAndExtract };
