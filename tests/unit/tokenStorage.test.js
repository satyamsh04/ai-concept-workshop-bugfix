describe('Token Storage', () => {
  const GATEWAY_TOKEN = '56a6491a7961caed0b413c43880cbc151181621bf0055047';

  beforeEach(() => {
    global.localStorage.getItem.mockClear();
    global.localStorage.setItem.mockClear();
  });

  const loadToken = () => {
    try {
      const stored = global.localStorage.getItem('acad-gateway-token');
      if (!stored) {
        global.localStorage.setItem('acad-gateway-token', GATEWAY_TOKEN);
      }
      return stored || GATEWAY_TOKEN;
    } catch (_) {
      return GATEWAY_TOKEN;
    }
  };

  const saveToken = (token) => {
    try {
      global.localStorage.setItem('acad-gateway-token', token);
      return true;
    } catch (_) {
      return false;
    }
  };

  describe('loadToken', () => {
    test('returns stored token when present', () => {
      global.localStorage.getItem.mockReturnValue('stored-token-123');
      expect(loadToken()).toBe('stored-token-123');
      expect(global.localStorage.getItem).toHaveBeenCalledWith('acad-gateway-token');
    });

    test('returns default token and stores it when nothing stored', () => {
      global.localStorage.getItem.mockReturnValue(null);
      const result = loadToken();
      expect(result).toBe(GATEWAY_TOKEN);
      expect(global.localStorage.setItem).toHaveBeenCalledWith('acad-gateway-token', GATEWAY_TOKEN);
    });

    test('returns default token if localStorage throws', () => {
      global.localStorage.getItem.mockImplementation(() => {
        throw new Error('Storage unavailable');
      });
      const result = loadToken();
      expect(result).toBe(GATEWAY_TOKEN);
    });

    test('handles empty string stored', () => {
      global.localStorage.getItem.mockReturnValue('');
      const result = loadToken();
      expect(result).toBe(GATEWAY_TOKEN);
    });
  });

  describe('saveToken', () => {
    test('successfully saves token', () => {
      global.localStorage.setItem.mockReturnValue(undefined);
      expect(saveToken('new-token')).toBe(true);
      expect(global.localStorage.setItem).toHaveBeenCalledWith('acad-gateway-token', 'new-token');
    });

    test('returns false when localStorage throws', () => {
      global.localStorage.setItem.mockImplementation(() => {
        throw new Error('Storage full');
      });
      expect(saveToken('new-token')).toBe(false);
    });
  });

  describe('token persistence', () => {
    test('token survives multiple loadToken calls', () => {
      global.localStorage.getItem.mockReturnValue('persistent-token');
      expect(loadToken()).toBe('persistent-token');
      expect(loadToken()).toBe('persistent-token');
      expect(global.localStorage.getItem).toHaveBeenCalledTimes(2);
    });

    test('user override replaces default token', () => {
      global.localStorage.getItem.mockReturnValue('user-custom-token');
      expect(loadToken()).toBe('user-custom-token');
      expect(loadToken()).not.toBe(GATEWAY_TOKEN);
    });
  });
});

describe('Custom Instructions Storage', () => {
  beforeEach(() => {
    global.localStorage.getItem.mockClear();
    global.localStorage.setItem.mockClear();
  });

  const getSavedSystemPrompt = () => {
    try {
      return global.localStorage.getItem('acad-custom-instructions') || '';
    } catch (_) {
      return '';
    }
  };

  const saveSystemPrompt = (prompt) => {
    try {
      global.localStorage.setItem('acad-custom-instructions', prompt);
      return true;
    } catch (_) {
      return false;
    }
  };

  test('getSavedSystemPrompt returns empty string when no prompt saved', () => {
    global.localStorage.getItem.mockReturnValue(null);
    expect(getSavedSystemPrompt()).toBe('');
  });

  test('getSavedSystemPrompt returns stored prompt', () => {
    global.localStorage.getItem.mockReturnValue('Always reply formally');
    expect(getSavedSystemPrompt()).toBe('Always reply formally');
  });

  test('saveSystemPrompt stores prompt', () => {
    global.localStorage.setItem.mockReturnValue(undefined);
    expect(saveSystemPrompt('Test prompt')).toBe(true);
    expect(global.localStorage.setItem).toHaveBeenCalledWith('acad-custom-instructions', 'Test prompt');
  });
});
