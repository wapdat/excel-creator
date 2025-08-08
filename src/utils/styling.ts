import { CellStyle, BorderStyle } from '../types';

export class StyleUtils {
  /**
   * Default header style
   */
  static getDefaultHeaderStyle(): CellStyle {
    return {
      font: {
        bold: true,
        size: 12,
        color: { argb: 'FF000000' }
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      },
      border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'medium', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      }
    };
  }

  /**
   * Default data cell style
   */
  static getDefaultDataStyle(): CellStyle {
    return {
      font: {
        size: 11,
        color: { argb: 'FF000000' }
      },
      border: {
        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
      },
      alignment: {
        horizontal: 'left',
        vertical: 'middle',
        wrapText: false
      }
    };
  }

  /**
   * Alternate row style
   */
  static getAlternateRowStyle(): CellStyle {
    return {
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF5F5F5' }
      }
    };
  }

  /**
   * Currency cell style
   */
  static getCurrencyStyle(symbol: string = '$'): CellStyle {
    return {
      numFmt: `"${symbol}"#,##0.00`,
      alignment: {
        horizontal: 'right',
        vertical: 'middle'
      }
    };
  }

  /**
   * Percentage cell style
   */
  static getPercentageStyle(decimals: number = 2): CellStyle {
    return {
      numFmt: decimals > 0 ? `0.${'0'.repeat(decimals)}%` : '0%',
      alignment: {
        horizontal: 'right',
        vertical: 'middle'
      }
    };
  }

  /**
   * Date cell style
   */
  static getDateStyle(format: string = 'yyyy-mm-dd'): CellStyle {
    return {
      numFmt: format,
      alignment: {
        horizontal: 'center',
        vertical: 'middle'
      }
    };
  }

  /**
   * Number cell style
   */
  static getNumberStyle(decimals: number = 0): CellStyle {
    return {
      numFmt: decimals > 0 ? `0.${'0'.repeat(decimals)}` : '0',
      alignment: {
        horizontal: 'right',
        vertical: 'middle'
      }
    };
  }

  /**
   * Total/summary row style
   */
  static getTotalRowStyle(): CellStyle {
    return {
      font: {
        bold: true,
        size: 11,
        color: { argb: 'FF000000' }
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD0E0F0' }
      },
      border: {
        top: { style: 'medium', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'double', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      },
      alignment: {
        horizontal: 'right',
        vertical: 'middle'
      }
    };
  }

  /**
   * Merge styles (override takes precedence)
   */
  static mergeStyles(base?: CellStyle, override?: CellStyle): CellStyle | undefined {
    if (!base && !override) return undefined;
    if (!base) return override;
    if (!override) return base;

    return {
      font: { ...base.font, ...override.font },
      fill: override.fill || base.fill,
      border: {
        top: override.border?.top || base.border?.top,
        left: override.border?.left || base.border?.left,
        bottom: override.border?.bottom || base.border?.bottom,
        right: override.border?.right || base.border?.right
      },
      alignment: { ...base.alignment, ...override.alignment },
      numFmt: override.numFmt || base.numFmt
    };
  }

  /**
   * Create style for specific column type
   */
  static getColumnStyle(format?: string, customStyle?: CellStyle): CellStyle | undefined {
    let baseStyle: CellStyle | undefined;

    switch (format) {
      case 'currency':
        baseStyle = this.getCurrencyStyle();
        break;
      case 'percentage':
        baseStyle = this.getPercentageStyle();
        break;
      case 'date':
        baseStyle = this.getDateStyle();
        break;
      case 'number':
        baseStyle = this.getNumberStyle(2);
        break;
      default:
        baseStyle = undefined;
    }

    return this.mergeStyles(baseStyle, customStyle);
  }

  /**
   * Get color palette for charts and conditional formatting
   */
  static getColorPalette(): string[] {
    return [
      'FF4472C4', // Blue
      'FFED7D31', // Orange
      'FFA5A5A5', // Gray
      'FFFFC000', // Yellow
      'FF5B9BD5', // Light Blue
      'FF70AD47', // Green
      'FF264478', // Dark Blue
      'FF9E480E', // Brown
      'FF636363', // Dark Gray
      'FF997300'  // Dark Yellow
    ];
  }

  /**
   * Get gradient colors for conditional formatting
   */
  static getGradientColors(type: 'green-red' | 'blue-white' | 'custom', customColors?: string[]): string[] {
    switch (type) {
      case 'green-red':
        return ['FF63BE7B', 'FFFFEB84', 'FFF8696B'];
      case 'blue-white':
        return ['FF5A8AC6', 'FFFFFF', 'FF5A8AC6'];
      case 'custom':
        return customColors || ['FFFFFFFF'];
      default:
        return ['FFFFFFFF'];
    }
  }

  /**
   * Create a custom border style
   */
  static createBorder(
    style: 'thin' | 'medium' | 'thick' | 'double' = 'thin',
    color: string = 'FF000000'
  ): BorderStyle {
    return {
      style,
      color: { argb: color }
    };
  }

  /**
   * Apply company branding colors
   */
  static getBrandedHeaderStyle(primaryColor: string, textColor: string = 'FFFFFFFF'): CellStyle {
    return {
      font: {
        bold: true,
        size: 12,
        color: { argb: textColor }
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: primaryColor }
      },
      border: {
        top: { style: 'thin', color: { argb: primaryColor } },
        left: { style: 'thin', color: { argb: primaryColor } },
        bottom: { style: 'medium', color: { argb: primaryColor } },
        right: { style: 'thin', color: { argb: primaryColor } }
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      }
    };
  }

  /**
   * Style presets
   */
  static getStylePreset(preset: 'minimal' | 'professional' | 'colorful' | 'dark'): {
    headerStyle: CellStyle;
    dataStyle: CellStyle;
    alternateRowStyle: CellStyle;
  } {
    switch (preset) {
      case 'minimal':
        return {
          headerStyle: {
            font: { bold: true, size: 11 },
            border: {
              bottom: { style: 'thin', color: { argb: 'FF000000' } }
            }
          },
          dataStyle: {},
          alternateRowStyle: {}
        };

      case 'professional':
        return {
          headerStyle: this.getDefaultHeaderStyle(),
          dataStyle: this.getDefaultDataStyle(),
          alternateRowStyle: this.getAlternateRowStyle()
        };

      case 'colorful':
        return {
          headerStyle: this.getBrandedHeaderStyle('FF4472C4', 'FFFFFFFF'),
          dataStyle: this.getDefaultDataStyle(),
          alternateRowStyle: {
            fill: {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE8F2FF' }
            }
          }
        };

      case 'dark':
        return {
          headerStyle: {
            font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } },
            fill: {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF2B2B2B' }
            },
            border: {
              top: { style: 'thin', color: { argb: 'FF505050' } },
              left: { style: 'thin', color: { argb: 'FF505050' } },
              bottom: { style: 'medium', color: { argb: 'FF505050' } },
              right: { style: 'thin', color: { argb: 'FF505050' } }
            },
            alignment: {
              horizontal: 'center',
              vertical: 'middle'
            }
          },
          dataStyle: {
            font: { size: 11, color: { argb: 'FFE0E0E0' } },
            fill: {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF1A1A1A' }
            },
            border: {
              top: { style: 'thin', color: { argb: 'FF404040' } },
              left: { style: 'thin', color: { argb: 'FF404040' } },
              bottom: { style: 'thin', color: { argb: 'FF404040' } },
              right: { style: 'thin', color: { argb: 'FF404040' } }
            }
          },
          alternateRowStyle: {
            fill: {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF252525' }
            }
          }
        };

      default:
        return {
          headerStyle: this.getDefaultHeaderStyle(),
          dataStyle: this.getDefaultDataStyle(),
          alternateRowStyle: this.getAlternateRowStyle()
        };
    }
  }
}