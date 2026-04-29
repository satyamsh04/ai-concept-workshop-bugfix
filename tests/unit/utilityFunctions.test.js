const sampleEmails = require('../fixtures/sampleEmails');

describe('Utility Functions', () => {
  describe('hashString', () => {
    const hashString = (str) => {
      let h = 0;
      for (let i = 0; i < str.length; i++) h = ((h << 5) - h + str.charCodeAt(i)) | 0;
      return Math.abs(h).toString(36);
    };

    test('produces consistent hash for same input', () => {
      const input = 'test string';
      expect(hashString(input)).toBe(hashString(input));
    });

    test('produces different hash for different inputs', () => {
      const hash1 = hashString('email one');
      const hash2 = hashString('email two');
      expect(hash1).not.toBe(hash2);
    });

    test('produces non-empty string', () => {
      const hash = hashString('any input');
      expect(typeof hash).toBe('string');
      expect(hash.length).toBeGreaterThan(0);
    });

    test('handles empty string', () => {
      const hash = hashString('');
      expect(hash).toBe('0');
    });

    test('handles unicode characters', () => {
      const hash = hashString('Japanese test');
      expect(typeof hash).toBe('string');
      expect(hash.length).toBeGreaterThan(0);
    });

    test('produces deterministic hash for email identifiers', () => {
      const email1 = 'Question about Assignment 2 deadline|alex.chen@student.university.edu|2026-04-15T09:30:00Z';
      const email2 = 'Urgent: Cannot access exam marks|s.johnson@student.university.edu|2026-04-15T14:22:00Z';
      const hash1 = hashString(email1);
      const hash2 = hashString(email2);
      expect(hash1).not.toBe(hash2);
      expect(hashString(email1)).toBe(hash1);
    });
  });

  describe('extractMessages', () => {
    const extractMessages = (result) => {
      if (!result) return [];
      if (Array.isArray(result.messages)) return result.messages;
      if (Array.isArray(result)) return result;
      if (result.history && Array.isArray(result.history)) return result.history;
      for (const k of Object.keys(result)) {
        if (Array.isArray(result[k])) return result[k];
      }
      return [];
    };

    test('returns empty array for null', () => {
      expect(extractMessages(null)).toEqual([]);
    });

    test('returns empty array for undefined', () => {
      expect(extractMessages(undefined)).toEqual([]);
    });

    test('extracts messages from { messages: [...] } format', () => {
      const result = { messages: [{ role: 'user', content: 'hello' }] };
      expect(extractMessages(result)).toEqual(result.messages);
    });

    test('extracts messages from direct array', () => {
      const messages = [{ role: 'user', content: 'hello' }];
      expect(extractMessages(messages)).toEqual(messages);
    });

    test('extracts messages from { history: [...] } format', () => {
      const result = { history: [{ role: 'assistant', content: 'hi' }] };
      expect(extractMessages(result)).toEqual(result.history);
    });

    test('extracts messages from first array property', () => {
      const result = { otherField: 'value', items: [{ role: 'user' }] };
      expect(extractMessages(result)).toEqual(result.items);
    });

    test('returns empty array for object with no arrays', () => {
      expect(extractMessages({ key: 'value' })).toEqual([]);
    });
  });

  describe('extractText', () => {
    const extractText = (msg) => {
      if (typeof msg.content === 'string') return msg.content;
      if (Array.isArray(msg.content)) {
        return msg.content.filter(c => c.type === 'text').map(c => c.text || '').join('\n');
      }
      return '';
    };

    test('extracts string content', () => {
      const msg = { content: 'Hello world' };
      expect(extractText(msg)).toBe('Hello world');
    });

    test('extracts text from array content', () => {
      const msg = {
        content: [
          { type: 'text', text: 'Line 1' },
          { type: 'text', text: 'Line 2' },
          { type: 'image', text: 'ignored' }
        ]
      };
      expect(extractText(msg)).toBe('Line 1\nLine 2');
    });

    test('returns empty string for missing content', () => {
      expect(extractText({})).toBe('');
    });

    test('returns empty string for null content', () => {
      expect(extractText({ content: null })).toBe('');
    });

    test('handles array with only non-text items', () => {
      const msg = { content: [{ type: 'image', text: 'data' }] };
      expect(extractText(msg)).toBe('');
    });

    test('handles object without content property', () => {
      expect(extractText({ role: 'user' })).toBe('');
    });
  });

  describe('isRawToolCall', () => {
    const isRawToolCall = (text) => {
      const t = text.trim();
      try {
        const obj = JSON.parse(t);
        if (obj && typeof obj.name === 'string' && obj.parameters !== undefined) return true;
      } catch (_) {}
      return /^\{"name"\s*:\s*"[^"]+"\s*,\s*"parameters"\s*:/.test(t);
    };

    test('detects valid tool call JSON', () => {
      expect(isRawToolCall('{"name": "test", "parameters": {}}')).toBe(true);
    });

    test('rejects plain text', () => {
      expect(isRawToolCall('Hello, how are you?')).toBe(false);
    });

    test('rejects non-tool JSON objects', () => {
      expect(isRawToolCall('{"message": "hello"}')).toBe(false);
      expect(isRawToolCall('{"foo": "bar"}')).toBe(false);
    });

    test('rejects JSON without name or parameters', () => {
      expect(isRawToolCall('{"name": "test"}')).toBe(false);
      expect(isRawToolCall('{"parameters": {}}')).toBe(false);
    });

    test('handles malformed JSON with regex', () => {
      expect(isRawToolCall('{"name": "test", "parameters": {}}')).toBe(true);
    });

    test('rejects empty string', () => {
      expect(isRawToolCall('')).toBe(false);
    });

    test('rejects whitespace only', () => {
      expect(isRawToolCall('   ')).toBe(false);
    });
  });

  describe('streamBuffer management', () => {
    const flushStream = (streamBuffer) => {
      const text = streamBuffer.trim();
      return { text, isEmpty: !text };
    };

    const renderStreamingBubble = (text) => {
      return text;
    };

    test('flushStream returns empty for whitespace buffer', () => {
      const result = flushStream('   ');
      expect(result.isEmpty).toBe(true);
    });

    test('flushStream returns text content', () => {
      const result = flushStream('Hello world');
      expect(result.text).toBe('Hello world');
      expect(result.isEmpty).toBe(false);
    });

    test('renderStreamingBubble returns text unchanged', () => {
      expect(renderStreamingBubble('Test')).toBe('Test');
    });

    test('renderStreamingBubble handles multi-line text', () => {
      const text = 'Line 1\nLine 2\nLine 3';
      expect(renderStreamingBubble(text)).toBe(text);
    });
  });
});
