import { useState } from 'react';
import type { CSSProperties } from 'react';
import { MaterialSymbol } from './MaterialSymbol';

interface VariablesSubmenuProps {
  variables: { [category: string]: string[] };
  onInsert: (token: string) => void;
  closeMenu: () => void;
}

const rowStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '6px 12px',
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  fontSize: 13,
  color: 'var(--doc-text, #374151)',
  width: '100%',
  textAlign: 'left',
  whiteSpace: 'nowrap',
  position: 'relative',
};

const subPanelStyle: CSSProperties = {
  position: 'absolute',
  left: '100%',
  top: -4,
  marginLeft: 2,
  backgroundColor: 'white',
  border: '1px solid var(--doc-border, #d1d5db)',
  borderRadius: 6,
  boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
  padding: '4px 0',
  zIndex: 1002,
  minWidth: 160,
};

export function VariablesSubmenu({ variables, onInsert, closeMenu }: VariablesSubmenuProps) {
  const [hoveredCategory, setHoveredCategory] = useState<string | null>(null);
  const categories = Object.keys(variables);

  return (
    <div style={{ padding: '4px 0', minWidth: 160 }}>
      {categories.map((category) => (
        <div
          key={category}
          style={{ position: 'relative' }}
          onMouseEnter={() => setHoveredCategory(category)}
          onMouseLeave={() => setHoveredCategory(null)}
        >
          <button
            type="button"
            style={rowStyle}
            onMouseDown={(e) => e.preventDefault()}
            onMouseOver={(e) => {
              (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                'var(--doc-hover, #f3f4f6)';
            }}
            onMouseOut={(e) => {
              (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
            }}
          >
            <span style={{ flex: 1 }}>{category}</span>
            <MaterialSymbol name="keyboard_arrow_right" size={16} />
          </button>
          {hoveredCategory === category && variables[category].length > 0 && (
            <div style={subPanelStyle} onMouseDown={(e) => e.preventDefault()}>
              {variables[category].map((varName) => {
                const token = `{{${category}.${varName}}}`;
                return (
                  <button
                    key={varName}
                    type="button"
                    style={{ ...rowStyle, position: 'static' }}
                    onClick={() => {
                      onInsert(token);
                      closeMenu();
                    }}
                    onMouseDown={(e) => e.preventDefault()}
                    onMouseOver={(e) => {
                      (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                        'var(--doc-hover, #f3f4f6)';
                    }}
                    onMouseOut={(e) => {
                      (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
                    }}
                  >
                    {varName}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}
