import * as React from 'react';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectGroup,
  SelectLabel,
  SelectSeparator,
  SelectTrigger,
  SelectValue,
} from './Select';
import { cn } from '../../lib/utils';
import { useTranslation } from '../../i18n';

// ============================================================================
// TYPES
// ============================================================================

export interface FontOption {
  name: string;
  fontFamily: string;
  category?: 'sans-serif' | 'serif' | 'monospace' | 'other';
}

export interface FontPickerProps {
  value?: string;
  onChange?: (fontFamily: string) => void;
  fonts?: FontOption[];
  /** Fonts to surface at the top (e.g. fonts used in the current document). */
  featuredFonts?: FontOption[];
  /**
   * Max fonts shown in the default (non-search) grouped view. Prevents slow
   * rendering when a large list is passed. Search always spans the full list.
   * Default: 150.
   */
  defaultLimit?: number;
  disabled?: boolean;
  className?: string;
  placeholder?: string;
  width?: number | string;
  showPreview?: boolean;
}

// ============================================================================
// DEFAULT FONTS
// ============================================================================

const DEFAULT_FONTS: FontOption[] = [
  { name: 'Arial', fontFamily: 'Arial, Helvetica, sans-serif', category: 'sans-serif' },
  { name: 'Calibri', fontFamily: '"Calibri", Arial, sans-serif', category: 'sans-serif' },
  { name: 'Helvetica', fontFamily: 'Helvetica, Arial, sans-serif', category: 'sans-serif' },
  { name: 'Verdana', fontFamily: 'Verdana, Geneva, sans-serif', category: 'sans-serif' },
  { name: 'Open Sans', fontFamily: '"Open Sans", sans-serif', category: 'sans-serif' },
  { name: 'Roboto', fontFamily: 'Roboto, sans-serif', category: 'sans-serif' },
  { name: 'Times New Roman', fontFamily: '"Times New Roman", Times, serif', category: 'serif' },
  { name: 'Georgia', fontFamily: 'Georgia, serif', category: 'serif' },
  { name: 'Cambria', fontFamily: 'Cambria, Georgia, serif', category: 'serif' },
  { name: 'Garamond', fontFamily: 'Garamond, serif', category: 'serif' },
  { name: 'Courier New', fontFamily: '"Courier New", Courier, monospace', category: 'monospace' },
  { name: 'Consolas', fontFamily: 'Consolas, monospace', category: 'monospace' },
];

// ============================================================================
// COMPONENT
// ============================================================================

export function FontPicker({
  value,
  onChange,
  fonts = DEFAULT_FONTS,
  featuredFonts,
  defaultLimit = 150,
  disabled = false,
  className,
  placeholder = 'Arial',
  width = 120,
  showPreview = true,
}: FontPickerProps) {
  const { t } = useTranslation();
  const [search, setSearch] = React.useState('');

  // Combined pool for display-name lookup (featured may not be in main list)
  const allFonts = React.useMemo(() => {
    const seen = new Set<string>();
    const out: FontOption[] = [];
    for (const f of [...(featuredFonts ?? []), ...fonts]) {
      if (!seen.has(f.name)) {
        seen.add(f.name);
        out.push(f);
      }
    }
    return out;
  }, [fonts, featuredFonts]);

  const displayValue = React.useMemo(() => {
    if (!value) return placeholder;
    const font = allFonts.find(
      (f) => f.fontFamily === value || f.name.toLowerCase() === value.toLowerCase()
    );
    return font?.name || value;
  }, [value, allFonts, placeholder]);

  const handleValueChange = React.useCallback(
    (newValue: string) => {
      const font = allFonts.find((f) => f.name === newValue);
      if (font) onChange?.(font.fontFamily);
    },
    [onChange, allFonts]
  );

  const handleOpenChange = React.useCallback((open: boolean) => {
    if (!open) setSearch('');
  }, []);

  // Build the visible lists based on search query
  const { visibleFeatured, groupedFonts, isSearching, hiddenCount } = React.useMemo(() => {
    const q = search.trim().toLowerCase();
    const isSearching = q.length > 0;

    const featuredNames = new Set((featuredFonts ?? []).map((f) => f.name));

    const visibleFeatured = (featuredFonts ?? []).filter(
      (f) => !isSearching || f.name.toLowerCase().includes(q)
    );

    const groups: Record<string, FontOption[]> = {
      'sans-serif': [],
      serif: [],
      monospace: [],
      other: [],
    };

    for (const font of fonts) {
      if (featuredNames.has(font.name)) continue; // already in featured section
      if (isSearching && !font.name.toLowerCase().includes(q)) continue;
      groups[font.category || 'other'].push(font);
    }

    for (const cat of Object.keys(groups)) {
      groups[cat].sort((a, b) => a.name.localeCompare(b.name));
    }

    // When not searching, cap total rendered fonts to keep the DOM fast.
    // Search always scans the full list so users can still find anything.
    let hiddenCount = 0;
    if (!isSearching && defaultLimit > 0) {
      let remaining = defaultLimit;
      for (const cat of Object.keys(groups)) {
        if (remaining <= 0) {
          hiddenCount += groups[cat].length;
          groups[cat] = [];
        } else if (groups[cat].length > remaining) {
          hiddenCount += groups[cat].length - remaining;
          groups[cat] = groups[cat].slice(0, remaining);
          remaining = 0;
        } else {
          remaining -= groups[cat].length;
        }
      }
    }

    return { visibleFeatured, groupedFonts: groups, isSearching, hiddenCount };
  }, [fonts, featuredFonts, search, defaultLimit]);

  const hasResults =
    visibleFeatured.length > 0 ||
    Object.values(groupedFonts).some((g) => g.length > 0);

  const renderItem = (font: FontOption, keyPrefix = '') => (
    <SelectItem
      key={`${keyPrefix}${font.name}`}
      value={font.name}
      style={showPreview ? { fontFamily: font.fontFamily } : undefined}
    >
      {font.name}
    </SelectItem>
  );

  // When searching, render a flat alphabetical list across all matching fonts
  const flatSearchResults = React.useMemo(() => {
    if (!isSearching) return [];
    return [...visibleFeatured, ...Object.values(groupedFonts).flat()].sort((a, b) =>
      a.name.localeCompare(b.name)
    );
  }, [isSearching, visibleFeatured, groupedFonts]);

  return (
    <Select
      value={displayValue}
      onValueChange={handleValueChange}
      disabled={disabled}
      onOpenChange={handleOpenChange}
    >
      <SelectTrigger
        className={cn('h-8 text-sm', className)}
        style={{ minWidth: typeof width === 'number' ? `${width}px` : width }}
        aria-label={t('font.selectAriaLabel')}
      >
        <SelectValue placeholder={placeholder}>{displayValue}</SelectValue>
      </SelectTrigger>
      <SelectContent className="max-h-[360px] overflow-hidden flex flex-col p-0">
        {/* Search input — stopPropagation prevents Radix typeahead from eating keystrokes */}
        <div className="px-2 py-1.5 border-b shrink-0">
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            onKeyDown={(e) => e.stopPropagation()}
            placeholder={t('font.searchPlaceholder') || 'Search fonts…'}
            className="w-full h-7 text-sm px-2 rounded border border-input bg-background outline-none"
          />
        </div>

        <div className="overflow-y-auto flex-1">
          {isSearching ? (
            // Flat alphabetical results when searching
            <SelectGroup>
              {flatSearchResults.map((f) => renderItem(f, 'search-'))}
            </SelectGroup>
          ) : (
            <>
              {/* Document fonts */}
              {visibleFeatured.length > 0 && (
                <SelectGroup>
                  <SelectLabel>{t('font.documentFonts') || 'Document'}</SelectLabel>
                  {visibleFeatured.map((f) => renderItem(f, 'featured-'))}
                </SelectGroup>
              )}

              {/* Category groups */}
              {groupedFonts['sans-serif'].length > 0 && (
                <>
                  {visibleFeatured.length > 0 && <SelectSeparator />}
                  <SelectGroup>
                    <SelectLabel>{t('font.sansSerif')}</SelectLabel>
                    {groupedFonts['sans-serif'].map((f) => renderItem(f))}
                  </SelectGroup>
                </>
              )}
              {groupedFonts['serif'].length > 0 && (
                <>
                  <SelectSeparator />
                  <SelectGroup>
                    <SelectLabel>{t('font.serif')}</SelectLabel>
                    {groupedFonts['serif'].map((f) => renderItem(f))}
                  </SelectGroup>
                </>
              )}
              {groupedFonts['monospace'].length > 0 && (
                <>
                  <SelectSeparator />
                  <SelectGroup>
                    <SelectLabel>{t('font.monospace')}</SelectLabel>
                    {groupedFonts['monospace'].map((f) => renderItem(f))}
                  </SelectGroup>
                </>
              )}
              {groupedFonts['other'].length > 0 && (
                <>
                  <SelectSeparator />
                  <SelectGroup>
                    {groupedFonts['other'].map((f) => renderItem(f))}
                  </SelectGroup>
                </>
              )}
            </>
          )}

          {!hasResults && (
            <div className="py-4 text-center text-sm text-muted-foreground">
              {t('font.noFontsFound') || 'No fonts found'}
            </div>
          )}

          {!isSearching && hiddenCount > 0 && (
            <div className="px-3 py-2 text-center text-xs text-muted-foreground border-t">
              {`Search to find ${hiddenCount.toLocaleString()} more fonts`}
            </div>
          )}
        </div>
      </SelectContent>
    </Select>
  );
}
