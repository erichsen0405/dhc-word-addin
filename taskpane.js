/* global Office, Word */

const BASE44_FUNCTION_URL = 'https://app.base44.com/api/v1/apps/691431d05cfac7d7acfaf766/functions/translateText';

let translatedText = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('translateBtn').addEventListener('click', translateDocument);
        document.getElementById('applyBtn').addEventListener('click', applyTranslation);
        document.getElementById('cancelBtn').addEventListener('click', cancelTranslation);
    }
});

async function translateDocument() {
    const targetLanguage = document.getElementById('targetLanguage').value;
    const translateSelectionOnly = document.getElementById('translateSelection').checked;
    
    if (!targetLanguage) {
        showStatus('Vælg venligst et sprog', 'error');
        return;
    }
    
    const btn = document.getElementById('translateBtn');
    btn.disabled = true;
    btn.querySelector('.btn-text').style.display = 'none';
    btn.querySelector('.btn-loading').style.display = 'inline';
    showStatus('Henter tekst...', 'info');
    
    try {
        await Word.run(async (context) => {
            let textToTranslate = '';
            
            if (translateSelectionOnly) {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                textToTranslate = selection.text;
                
                if (!textToTranslate?.trim()) {
                    throw new Error('Marker venligst tekst først');
                }
            } else {
                const body = context.document.body;
                body.load('text');
                await context.sync();
                textToTranslate = body.text;
            }
            
            showStatus('Oversætter...', 'info');
            
            const response = await fetch(BASE44_FUNCTION_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ text: textToTranslate, targetLanguage })
            });
            
            if (!response.ok) throw new Error('Oversættelse fejlede');
            
            const result = await response.json();
            translatedText = result.translatedText;
            
            document.getElementById('previewContent').textContent = 
                translatedText.substring(0, 500) + (translatedText.length > 500 ? '...' : '');
            document.getElementById('preview').style.display = 'block';
            showStatus('Oversættelse klar!', 'success');
        });
    } catch (error) {
        showStatus('Fejl: ' + error.message, 'error');
    } finally {
        btn.disabled = false;
        btn.querySelector('.btn-text').style.display = 'inline';
        btn.querySelector('.btn-loading').style.display = 'none';
    }
}

async function applyTranslation() {
    if (!translatedText) return;
    
    try {
        await Word.run(async (context) => {
            if (document.getElementById('translateSelection').checked) {
                context.document.getSelection().insertText(translatedText, Word.InsertLocation.replace);
            } else {
                context.document.body.clear();
                context.document.body.insertText(translatedText, Word.InsertLocation.start);
            }
            await context.sync();
            
            showStatus('Indsat!', 'success');
            document.getElementById('preview').style.display = 'none';
            translatedText = '';
        });
    } catch (error) {
        showStatus('Fejl: ' + error.message, 'error');
    }
}

function cancelTranslation() {
    translatedText = '';
    document.getElementById('preview').style.display = 'none';
    showStatus('', '');
}

function showStatus(message, type) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = 'status ' + type;
    status.style.display = message ? 'block' : 'none';
}