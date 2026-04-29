describe('Theme Detection', () => {
  const originalOfficeTheme = global.Office.Context.officeTheme;

  const applyTheme = () => {
    let dark = false;
    try {
      const t = global.Office.Context.officeTheme;
      if (t && t.bodyBackgroundColor) {
        const hex = t.bodyBackgroundColor.replace('#', '');
        const r = parseInt(hex.slice(0, 2), 16);
        const g = parseInt(hex.slice(2, 4), 16);
        const b = parseInt(hex.slice(4, 6), 16);
        dark = (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
      }
    } catch (_) {}

    return dark ? 'dark' : 'light';
  };

  beforeEach(() => {
    global.Office.Context.officeTheme = null;
  });

  afterAll(() => {
    global.Office.Context.officeTheme = originalOfficeTheme;
  });

  describe('applyTheme with Office theme', () => {
    test('detects dark theme from officeTheme bodyBackgroundColor', () => {
      global.Office.Context.officeTheme = {
        bodyBackgroundColor: '#1E1E1E'
      };
      expect(applyTheme()).toBe('dark');
    });

    test('detects light theme from officeTheme bodyBackgroundColor', () => {
      global.Office.Context.officeTheme = {
        bodyBackgroundColor: '#FFFFFF'
      };
      expect(applyTheme()).toBe('light');
    });

    test('handles officeTheme with mid-grey color', () => {
      global.Office.Context.officeTheme = {
        bodyBackgroundColor: '#808080'
      };
      const luminance = (0.299 * 128 + 0.587 * 128 + 0.114 * 128) / 255;
      expect(luminance).toBeCloseTo(0.5, 2);
      expect(applyTheme()).toBe('light');
    });

    test('handles officeTheme with warm color', () => {
      global.Office.Context.officeTheme = {
        bodyBackgroundColor: '#FFF8F0'
      };
      expect(applyTheme()).toBe('light');
    });

    test('handles officeTheme without bodyBackgroundColor', () => {
      global.Office.Context.officeTheme = {};
      expect(applyTheme()).toBe('light');
    });

    test('handles null officeTheme', () => {
      global.Office.Context.officeTheme = null;
      expect(applyTheme()).toBe('light');
    });
  });

  describe('color luminance calculation', () => {
    test('dark colors produce dark theme', () => {
      const darkColors = ['#000000', '#1a1a1a', '#333333', '#0D1117', '#1E1E1E'];
      for (const color of darkColors) {
        global.Office.Context.officeTheme = { bodyBackgroundColor: color };
        expect(applyTheme()).toBe('dark');
      }
    });

    test('light colors produce light theme', () => {
      const lightColors = ['#FFFFFF', '#FAFAFA', '#F5F5F5', '#FFFFFE'];
      for (const color of lightColors) {
        global.Office.Context.officeTheme = { bodyBackgroundColor: color };
        expect(applyTheme()).toBe('light');
      }
    });
  });
});
