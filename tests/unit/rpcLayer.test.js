describe('RPC Layer', () => {
  let mockSocket;
  let rpcSeq;

  const createRpcModule = () => {
    let socket = mockSocket;
    let pendingRpc = new Map();
    let rpcSeq = 0;

    const callRpc = (method, params) => {
      return new Promise((resolve, reject) => {
        if (!socket || socket.readyState !== 1) {
          reject(new Error('Not connected'));
          return;
        }
        const id = String(++rpcSeq);
        pendingRpc.set(id, { resolve, reject });
        socket.send(JSON.stringify({ type: 'req', id, method, params }));
      });
    };

    const handleIncoming = (raw) => {
      let data;
      try {
        data = JSON.parse(raw);
      } catch (_) {
        return;
      }

      if (data.type === 'res' && pendingRpc.has(String(data.id))) {
        const { resolve, reject } = pendingRpc.get(String(data.id));
        pendingRpc.delete(String(data.id));
        if (data.ok === false) {
          reject(new Error(data.error?.message || data.error?.code || 'RPC error'));
        } else {
          resolve(data.result || data.payload || data);
        }
      }
    };

    return { callRpc, handleIncoming, getPendingCount: () => pendingRpc.size };
  };

  beforeEach(() => {
    pendingRpc = new Map();
    rpcSeq = 0;
    mockSocket = {
      readyState: 1,
      send: jest.fn()
    };
  });

  describe('callRpc', () => {
    test('creates pending RPC and sends message', async () => {
      const rpc = createRpcModule();
      const promise = rpc.callRpc('chat.send', { message: 'test' });
      expect(mockSocket.send).toHaveBeenCalled();
      const sentData = JSON.parse(mockSocket.send.mock.calls[0][0]);
      expect(sentData.type).toBe('req');
      expect(sentData.method).toBe('chat.send');
      expect(sentData.params).toEqual({ message: 'test' });
    });

    test('rejects when socket closed', async () => {
      mockSocket.readyState = 3;
      const rpc = createRpcModule();
      try {
        await rpc.callRpc('test', {});
        expect(false).toBe(true);
      } catch (e) {
        expect(e.message).toBe('Not connected');
      }
    });

    test('rejects when socket connecting', async () => {
      mockSocket.readyState = 0;
      const rpc = createRpcModule();
      try {
        await rpc.callRpc('test', {});
        expect(false).toBe(true);
      } catch (e) {
        expect(e.message).toBe('Not connected');
      }
    });
  });

  describe('handleIncoming', () => {
    test('resolves pending RPC on success response', async () => {
      const rpc = createRpcModule();
      const promise = rpc.callRpc('test', {});
      const sentData = JSON.parse(mockSocket.send.mock.calls[0][0]);

      rpc.handleIncoming(JSON.stringify({
        type: 'res',
        id: sentData.id,
        ok: true,
        result: { success: true }
      }));

      const result = await promise;
      expect(result).toEqual({ success: true });
    });

    test('rejects pending RPC on error response', async () => {
      const rpc = createRpcModule();
      const promise = rpc.callRpc('test', {});
      const sentData = JSON.parse(mockSocket.send.mock.calls[0][0]);

      rpc.handleIncoming(JSON.stringify({
        type: 'res',
        id: sentData.id,
        ok: false,
        error: { message: 'Test error' }
      }));

      try {
        await promise;
        expect(false).toBe(true);
      } catch (e) {
        expect(e.message).toBe('Test error');
      }
    });

    test('ignores response for unknown id', () => {
      const rpc = createRpcModule();
      expect(() => rpc.handleIncoming(JSON.stringify({
        type: 'res',
        id: '999',
        ok: true,
        result: {}
      }))).not.toThrow();
    });

    test('ignores malformed JSON', () => {
      const rpc = createRpcModule();
      expect(() => rpc.handleIncoming('not valid json')).not.toThrow();
    });
  });

  describe('RPC message format', () => {
    test('generates sequential RPC IDs', async () => {
      const rpc = createRpcModule();
      rpc.callRpc('method1', {});
      rpc.callRpc('method2', {});
      const id1 = JSON.parse(mockSocket.send.mock.calls[0][0]).id;
      const id2 = JSON.parse(mockSocket.send.mock.calls[1][0]).id;
      expect(id1).not.toBe(id2);
    });
  });
});

describe('Message History', () => {
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

  const extractText = (msg) => {
    if (typeof msg.content === 'string') return msg.content;
    if (Array.isArray(msg.content)) {
      return msg.content.filter(c => c.type === 'text').map(c => c.text || '').join('\n');
    }
    return '';
  };

  test('extracts user messages from history', () => {
    const history = {
      messages: [
        { role: 'user', content: 'What is this email about?' },
        { role: 'assistant', content: 'This email is an inquiry about assignment deadlines.' }
      ]
    };
    const messages = extractMessages(history);
    expect(messages).toHaveLength(2);
    expect(extractText(messages[0])).toBe('What is this email about?');
  });

  test('extracts assistant messages from history', () => {
    const history = {
      messages: [
        { role: 'user', content: 'Hello' },
        { role: 'assistant', content: 'Hi, how can I help you?' }
      ]
    };
    const messages = extractMessages(history);
    expect(extractText(messages[1])).toBe('Hi, how can I help you?');
  });

  test('handles history with openclaw metadata', () => {
    const history = {
      messages: [
        {
          role: 'assistant',
          content: 'Test response',
          __openclaw: { id: 'msg-123' }
        }
      ]
    };
    const messages = extractMessages(history);
    expect(messages[0].__openclaw.id).toBe('msg-123');
  });
});
