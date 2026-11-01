/**
 * Voice transcription used by the voice recorder: an Azure OpenAI / Foundry audio transcriptions endpoint
 * (e.g. gpt-4o-transcribe) configured per agent, called with the user's Entra token (audience https://ai.azure.com).
 */
export function createTranscriber(
  getToken: (resource: string) => Promise<string>,
  transcriptionUrl: string,
  transcriptionPrompt?: string
): (audio: Blob, lang: string) => Promise<string> {
  // Vocabulary hint: names in the prompt are spelled correctly in the transcript.
  const prompt = transcriptionPrompt?.trim();
  return async (audio, lang) => {
    const token = await getToken('https://ai.azure.com');
    const form = new FormData();
    form.append('file', audio, `dictation.${audio.type.includes('ogg') ? 'ogg' : audio.type.includes('mp4') ? 'mp4' : 'webm'}`);
    form.append('language', lang.slice(0, 2).toLowerCase());
    if (prompt) form.append('prompt', prompt);
    const res = await fetch(transcriptionUrl, { method: 'POST', headers: { Authorization: `Bearer ${token}` }, body: form });
    if (!res.ok) throw new Error(`Transcription ${res.status}`);
    const body = (await res.json()) as { text?: string };
    return body.text ?? '';
  };
}
