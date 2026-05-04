import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';

const PLOTTER_UNITS_PER_MM = 40; // Standard HPGL uses 40 plotter units per mm
const SECTION_GAP = 500; // 50cm gap between sections

interface CornerRadii {
    tl?: number;
    tr?: number;
    br?: number;
    bl?: number;
}

interface AppItem {
    id: string;
    internalOrderNumber: string;
    productCode: string;
    notes: string;
    material: string;
    color: string;
    w: number;
    h: number;
    qty: number;
    cornerRadius: CornerRadii;
    isCustom: boolean;
    classification: string;
    rotated?: boolean;
    rotationAngle?: number; // 0, 90, 180, 270
    x?: number;
    y?: number;
    originalId?: string;
    pathData?: string; // To store SVG path for irregular shapes
    isSupplement?: boolean; // New flag for copied items
    pageIndex?: number; // Track which page the item is on
}

interface PageLayout {
    pageIndex: number;
    boxes: AppItem[];
    w: number;
    h: number; // Actual used height
}

interface FullLayout {
    pages: PageLayout[];
    totalW: number;
    totalH: number; // Sum of used heights of all pages
}

const EXPORT_COLUMNS = [
    { key: 'index', label: '序号' },
    { key: 'internalOrderNumber', label: '内部订单号' },
    { key: 'productCode', label: '商品编码' },
    { key: 'notes', label: '卖家备注' },
    { key: 'material', label: '材质' },
    { key: 'color', label: '颜色' },
    { key: 'w', label: '宽(mm)' },
    { key: 'h', label: '高(mm)' },
    { key: 'rotated', label: '旋转' },
    { key: 'cost', label: '成本价' },
    { key: 'page', label: '分段' }
];

const PRESET_SUPPLEMENTS: Partial<AppItem>[] = [
    { w: 200, h: 250, cornerRadius: { tl: 10, tr: 10, br: 10, bl: 10 }, internalOrderNumber: '预设 200x250' },
    { w: 300, h: 450, cornerRadius: { tl: 150, tr: 150, br: 150, bl: 150 }, internalOrderNumber: '预设 300x450' },
    { w: 400, h: 800, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 400x800' },
    { w: 400, h: 1000, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 400x1000' },
    { w: 400, h: 1200, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 400x1200' },
    { w: 500, h: 800, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 500x800' },
    { w: 500, h: 1000, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 500x1000' },
    { w: 500, h: 1200, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 500x1200' },
    { w: 600, h: 1200, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 600x1200' },
    { w: 800, h: 1200, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 800x1200' },
    { w: 800, h: 1400, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, internalOrderNumber: '预设 800x1400' },
];

// --- Helper for Consistent Colors ---
const stringToColor = (str: string) => {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const c = (hash & 0x00FFFFFF).toString(16).toUpperCase();
    return '#' + '00000'.substring(0, 6 - c.length) + c;
};

// --- Global Best Fit MaxRects Logic (Material Saving) ---

interface FreeRect {
    x: number;
    y: number;
    w: number;
    h: number;
}

const MAX_INT = Number.MAX_SAFE_INTEGER;

// Strategy Definitions
type AlgorithmType = 'MAXRECTS' | 'SKYLINE' | 'SHELF';
type SortStrategy = 'AREA_DESC' | 'LONGSIDE_DESC' | 'SHORTSIDE_DESC' | 'PERIMETER_DESC' | 'WIDTH_DESC' | 'COMPLEMENTARY' | 'SMART_WIDTH_MATCH' | 'RANDOM' | 'ORDER_GROUP' | 'ORDER_PRIORITY_OPTIMIZED' | 'AUTO_MULTI_PRIORITY' | 'CUSTOM';
type Heuristic = 'BSSF' | 'BLSF' | 'BAF'; // Best Short Side Fit, Best Long Side Fit, Best Area Fit

interface SavedStrategy {
    id: string;
    name: string;
    algorithm: AlgorithmType;
    sortStrategy: SortStrategy;
    heuristic: Heuristic;
    allowRotation: boolean;
    utilization: number; 
    usageCount: number;
    createdAt: number;
    isCustom?: boolean;
    description?: string;
}

const STRATEGY_LIB_KEY = 'nesting_strategy_library';

const defaultStrategies: SavedStrategy[] = [
    { id: 'def-1', name: '经典面积降序', algorithm: 'MAXRECTS', sortStrategy: 'AREA_DESC', heuristic: 'BSSF', allowRotation: true, utilization: 0.85, usageCount: 0, createdAt: Date.now() },
    { id: 'def-2', name: '极速天际线', algorithm: 'SKYLINE', sortStrategy: 'WIDTH_DESC', heuristic: 'BSSF', allowRotation: true, utilization: 0.82, usageCount: 0, createdAt: Date.now() }
];

const loadStrategyLibrary = (): SavedStrategy[] => {
    const saved = localStorage.getItem(STRATEGY_LIB_KEY);
    if (!saved) return defaultStrategies;
    try {
        return JSON.parse(saved);
    } catch {
        return defaultStrategies;
    }
};

const saveStrategyLibrary = (lib: SavedStrategy[]) => {
    localStorage.setItem(STRATEGY_LIB_KEY, JSON.stringify(lib));
};

// Main entry for multi-page packing
const packLayout = (
    itemsInput: AppItem[], 
    containerWidth: number, 
    maxContainerHeight: number, // 0 means infinite
    spacing: number,
    allowRotation: boolean,
    useRandom: boolean,
    sortStrategy: SortStrategy = 'AREA_DESC',
    heuristic: Heuristic = 'BSSF',
    groupByOrder: boolean = false,
    algorithm: AlgorithmType = 'MAXRECTS'
): FullLayout => {
    
    // 1. Preparation & Sorting
    const items = itemsInput.map(i => ({...i}));

    // Strategy Implementation
    if (sortStrategy === 'CUSTOM') {
        // Do not sort, preserve input order (used for Genetic Algorithm sequences)
    } 
    else if (sortStrategy === 'AUTO_MULTI_PRIORITY') {
        const counts = new Map<string, number>();
        items.forEach(it => counts.set(it.internalOrderNumber, (counts.get(it.internalOrderNumber) || 0) + 1));
        const singles = items.filter(it => counts.get(it.internalOrderNumber) === 1);
        const multis = items.filter(it => counts.get(it.internalOrderNumber)! > 1);
        singles.sort((a, b) => (b.w * b.h) - (a.w * a.h));
        multis.sort((a, b) => {
            const orderComp = a.internalOrderNumber.localeCompare(b.internalOrderNumber, 'zh-CN', { numeric: true });
            if (orderComp !== 0) return orderComp;
            return (b.w * b.h) - (a.w * a.h);
        });
        items.splice(0, items.length, ...singles, ...multis);
    } 
    else if (sortStrategy === 'ORDER_PRIORITY_OPTIMIZED') {
        const orderGroups = new Map<string, AppItem[]>();
        items.forEach(item => {
            const list = orderGroups.get(item.internalOrderNumber) || [];
            list.push(item);
            orderGroups.set(item.internalOrderNumber, list);
        });
        const sortedOrderKeys = Array.from(orderGroups.keys()).sort((a, b) => {
             return a.localeCompare(b, 'zh-CN', { numeric: true });
        });
        const reorderedItems: AppItem[] = [];
        sortedOrderKeys.forEach(key => {
            const group = orderGroups.get(key);
            const tier0 = group.filter(i => Math.max(i.w, i.h) > 1400);
            const tier1 = group.filter(i => Math.max(i.w, i.h) <= 1400);
            tier0.sort((a, b) => Math.max(b.w, b.h) - Math.max(a.w, a.h));
            tier1.sort((a, b) => (b.w * b.h) - (a.w * a.h));
            reorderedItems.push(...tier0, ...tier1);
        });
        items.splice(0, items.length, ...reorderedItems);
    } 
    else if (groupByOrder) {
        items.sort((a, b) => {
            const orderComp = a.internalOrderNumber.localeCompare(b.internalOrderNumber, 'zh-CN', { numeric: true });
            if (orderComp !== 0) return orderComp;
            return (b.w * b.h) - (a.w * a.h);
        });
    } else {
        switch (sortStrategy) {
            case 'AREA_DESC': items.sort((a, b) => (b.w * b.h) - (a.w * a.h)); break;
            case 'LONGSIDE_DESC': items.sort((a, b) => Math.max(b.w, b.h) - Math.max(a.w, a.h)); break;
            case 'SHORTSIDE_DESC': items.sort((a, b) => Math.min(b.w, b.h) - Math.min(a.w, a.h)); break;
            case 'WIDTH_DESC':
                items.sort((a, b) => Math.max(b.w, b.h) - Math.max(a.w, a.h));
                items.sort((a, b) => b.w - a.w);
                break;
            case 'PERIMETER_DESC': items.sort((a, b) => (b.w + b.h) - (a.w + a.h)); break;
            case 'COMPLEMENTARY':
                const threshold = 1200;
                const longItems = items.filter(i => Math.max(i.w, i.h) >= threshold).sort((a,b) => (b.w * b.h) - (a.w * a.h));
                const shortItems = items.filter(i => Math.max(i.w, i.h) < threshold).sort((a,b) => (b.w * b.h) - (a.w * a.h));
                items.length = 0;
                while(longItems.length > 0 || shortItems.length > 0) {
                    if(longItems.length) items.push(longItems.shift());
                    if(shortItems.length) items.push(shortItems.shift());
                    if(shortItems.length) items.push(shortItems.shift());
                }
                break;
            case 'SMART_WIDTH_MATCH':
                const pool = [...items];
                pool.sort((a, b) => Math.max(b.w, b.h) - Math.max(a.w, a.h));
                const matchedItems = [];
                while(pool.length > 0) {
                    const current = pool.shift();
                    let bestConfig = { swapA: false, itemBIndex: -1, swapB: false, waste: MAX_INT };
                    const checkA = (widthA, swapA) => {
                         const remW = containerWidth - widthA - spacing;
                         if (remW < 0) return;
                         let bestIdx = -1; let swapB = false; let minRem = MAX_INT;
                         for(let i=0; i<pool.length; i++) {
                             const p = pool[i];
                             if (p.w <= remW) { const diff = remW - p.w; if (diff < minRem) { minRem = diff; bestIdx = i; swapB = false; } }
                             if (p.h <= remW) { const diff = remW - p.h; if (diff < minRem) { minRem = diff; bestIdx = i; swapB = true; } }
                             if (minRem === 0) break;
                         }
                         const waste = (bestIdx === -1) ? remW : minRem;
                         if (waste < bestConfig.waste) bestConfig = { swapA, itemBIndex: bestIdx, swapB, waste };
                    };
                    checkA(current.w, false); checkA(current.h, true);
                    if (bestConfig.swapA) { const temp = current.w; current.w = current.h; current.h = temp; const tR = current.cornerRadius; current.cornerRadius = { tl: tR.bl, tr: tR.tl, br: tR.tr, bl: tR.br }; }
                    matchedItems.push(current);
                    if (bestConfig.itemBIndex !== -1) {
                        const partner = pool[bestConfig.itemBIndex];
                        pool.splice(bestConfig.itemBIndex, 1);
                        if (bestConfig.swapB) { const temp = partner.w; partner.w = partner.h; partner.h = temp; const tR = partner.cornerRadius; partner.cornerRadius = { tl: tR.bl, tr: tR.tl, br: tR.tr, bl: tR.br }; }
                        matchedItems.push(partner);
                    }
                }
                items.splice(0, items.length, ...matchedItems);
                break;
            case 'RANDOM': items.sort(() => Math.random() - 0.5); break;
            default: items.sort((a, b) => (b.w * b.h) - (a.w * a.h));
        }
    }

    const pages: PageLayout[] = [];
    let remainingItems = items;
    let pageIndex = 1;

    // 2. Pagination Loop (Sections)
    while (remainingItems.length > 0) {
        const useSequentialPacking = groupByOrder || sortStrategy === 'ORDER_PRIORITY_OPTIMIZED' || sortStrategy === 'AUTO_MULTI_PRIORITY' || sortStrategy === 'CUSTOM';
        const result = packOnePage(remainingItems, containerWidth, maxContainerHeight, spacing, allowRotation, useRandom, heuristic, useSequentialPacking, pageIndex, sortStrategy);
        pages.push(result.page);
        remainingItems = result.unpacked;
        pageIndex++;
        if (result.page.boxes.length === 0 && remainingItems.length > 0) {
            console.warn("Item too large to fit in page constraints:", remainingItems[0]);
            remainingItems.shift(); 
        }
    }

    const totalH = pages.reduce((sum, p) => sum + p.h, 0);

    return { pages, totalW: containerWidth, totalH };
};

const packShelf = (
    itemsSource: AppItem[], 
    containerWidth: number, 
    maxH: number, 
    spacing: number, 
    allowRotation: boolean, 
    pageIndex: number
) => {
    const items = [...itemsSource];
    const packed: AppItem[] = [];
    const unpacked: AppItem[] = [];
    
    let currentX = 0;
    let currentY = 0;
    let shelfHeight = 0;

    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        let w = item.w;
        let h = item.h;
        let rotated = false;

        if (allowRotation && item.h > item.w && item.h <= containerWidth) {
            [w, h] = [h, w];
            rotated = true;
        }

        if (currentX + w > containerWidth) {
            currentX = 0;
            currentY += shelfHeight;
            shelfHeight = 0;
        }

        if (currentY + h > maxH) {
            unpacked.push(...items.slice(i));
            break;
        }

        packed.push({
            ...item,
            x: currentX,
            y: currentY,
            rotated,
            rotationAngle: rotated ? 90 : 0,
            pageIndex
        });

        currentX += w + spacing;
        shelfHeight = Math.max(shelfHeight, h + spacing);
    }

    const pageH = currentY + shelfHeight;
    return {
        page: { pageIndex, boxes: packed, w: containerWidth, h: pageH },
        unpacked
    };
};

interface SkylineSegment { x: number; y: number; w: number; }

const packSkyline = (
    itemsSource: AppItem[], 
    containerWidth: number, 
    maxH: number, 
    spacing: number, 
    allowRotation: boolean, 
    useRandom: boolean, 
    pageIndex: number
) => {
    const items = [...itemsSource];
    const packed: AppItem[] = [];
    const unpacked: AppItem[] = [];
    
    let skyline: SkylineSegment[] = [{ x: 0, y: 0, w: containerWidth }];

    const findLowestY = () => {
        let minY = MAX_INT;
        let index = -1;
        for (let i = 0; i < skyline.length; i++) {
            if (skyline[i].y < minY) {
                minY = skyline[i].y;
                index = i;
            }
        }
        return index;
    };

    while (items.length > 0) {
        let bestItemIdx = -1;
        let bestSegmentIdx = -1;
        let bestRotated = false;
        let bestY = MAX_INT;
        let bestRect = { x: 0, y: 0, w: 0, h: 0 };

        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const dims = [{ w: item.w, h: item.h, rot: false }];
            if (allowRotation) dims.push({ w: item.h, h: item.w, rot: true });

            for (const d of dims) {
                // Try to fit at each segment
                for (let j = 0; j < skyline.length; j++) {
                    if (skyline[j].x + d.w > containerWidth) continue;

                    // Calculate y at this position (max y of segments covered)
                    let currentY = 0;
                    let widthLeft = d.w;
                    for (let k = j; k < skyline.length && widthLeft > 0; k++) {
                        currentY = Math.max(currentY, skyline[k].y);
                        widthLeft -= skyline[k].w;
                    }

                    if (currentY + d.h > maxH) continue;

                    // Heuristic: Prefer lowest y, then lowest x
                    if (currentY < bestY) {
                        bestY = currentY;
                        bestItemIdx = i;
                        bestSegmentIdx = j;
                        bestRotated = d.rot;
                        bestRect = { x: skyline[j].x, y: currentY, w: d.w, h: d.h };
                    }
                }
            }
        }

        if (bestItemIdx === -1) break;

        const item = items[bestItemIdx];
        const newBox = { 
            ...item, 
            x: bestRect.x, 
            y: bestRect.y, 
            rotated: bestRotated, 
            rotationAngle: bestRotated ? 90 : 0, 
            pageIndex 
        };
        packed.push(newBox);
        items.splice(bestItemIdx, 1);

        // Update skyline
        const newSegment = { x: bestRect.x, y: bestRect.y + bestRect.h + spacing, w: bestRect.w + spacing };
        
        // Find segments covered
        let firstCovered = bestSegmentIdx;
        let lastCovered = bestSegmentIdx;
        let widthRemaining = bestRect.w;
        let k = bestSegmentIdx;
        while (k < skyline.length && widthRemaining > 0) {
            lastCovered = k;
            widthRemaining -= skyline[k].w;
            k++;
        }

        // Handle splitting last segment if not fully covered
        const lastSeg = skyline[lastCovered];
        const lastSegRight = lastSeg.x + lastSeg.w;
        const newRectRight = bestRect.x + bestRect.w;

        const replacements: SkylineSegment[] = [newSegment];
        if (newRectRight < lastSegRight) {
            replacements.push({ x: newRectRight, y: lastSeg.y, w: lastSegRight - newRectRight });
        }

        skyline.splice(firstCovered, lastCovered - firstCovered + 1, ...replacements);

        // Merge adjacent segments with same y
        for (let s = 0; s < skyline.length - 1; s++) {
            if (skyline[s].y === skyline[s+1].y) {
                skyline[s].w += skyline[s+1].w;
                skyline.splice(s + 1, 1);
                s--;
            }
        }
    }

    unpacked.push(...items);
    const pageH = packed.length > 0 ? Math.max(...packed.map(b => b.y + (b.rotated ? b.w : b.h))) : 0;

    return {
        page: { pageIndex, boxes: packed, w: containerWidth, h: pageH },
        unpacked
    };
};

// Helper: Pack a single page, return packed items and remaining items
const packOnePage = (
    itemsSource: AppItem[], 
    containerWidth: number, 
    maxHeight: number, 
    spacing: number, 
    allowRotation: boolean, 
    useRandom: boolean, 
    heuristic: Heuristic, 
    useSequentialPacking: boolean, 
    pageIndex: number, 
    sortStrategy?: SortStrategy,
    algorithm: AlgorithmType = 'MAXRECTS'
) => {
    const effectiveMaxH = (maxHeight > 0) ? maxHeight : MAX_INT;
    
    if (algorithm === 'SKYLINE') {
        return packSkyline(itemsSource, containerWidth, effectiveMaxH, spacing, allowRotation, useRandom, pageIndex);
    }
    if (algorithm === 'SHELF') {
        return packShelf(itemsSource, containerWidth, effectiveMaxH, spacing, allowRotation, pageIndex);
    }

    let freeRects: FreeRect[] = [{ x: 0, y: 0, w: containerWidth, h: effectiveMaxH }];
    
    const packed: AppItem[] = [];
    const unpacked: AppItem[] = [];
    const itemsToProcess = [...itemsSource];

    const isIntersecting = (r1: FreeRect, r2: {x:number, y:number, w:number, h:number}) => { return !(r2.x >= r1.x + r1.w || r2.x + r2.w <= r1.x || r2.y >= r1.y + r1.h || r2.y + r2.h <= r1.y); };
    const contains = (outer: FreeRect, inner: FreeRect) => { return inner.x >= outer.x && inner.y >= outer.y && inner.x + inner.w <= outer.x + outer.w && inner.y + inner.h <= outer.y + outer.h; };
    const pruneFreeRects = () => {
        for(let i = 0; i < freeRects.length; i++) {
            for(let j = i + 1; j < freeRects.length; j++) {
                if (contains(freeRects[j], freeRects[i])) { freeRects.splice(i, 1); i--; break; }
                if (contains(freeRects[i], freeRects[j])) { freeRects.splice(j, 1); j--; }
            }
        }
    };
    const placeItem = (item: AppItem, rect: FreeRect, rotated: boolean) => {
        const itemW = rotated ? item.h : item.w;
        const itemH = rotated ? item.w : item.h;
        const placedBox = { ...item, x: rect.x, y: rect.y, rotated, rotationAngle: rotated ? 90 : 0, pageIndex };
        packed.push(placedBox);
        const placedRect = { x: placedBox.x, y: placedBox.y, w: itemW + spacing, h: itemH + spacing };
        const newRects: FreeRect[] = [];
        for (let i = freeRects.length - 1; i >= 0; i--) {
            const r = freeRects[i];
            if (isIntersecting(r, placedRect)) {
                freeRects.splice(i, 1);
                if (placedRect.y > r.y && placedRect.y < r.y + r.h) newRects.push({ x: r.x, y: r.y, w: r.w, h: placedRect.y - r.y });
                if (placedRect.y + placedRect.h < r.y + r.h) newRects.push({ x: r.x, y: placedRect.y + placedRect.h, w: r.w, h: r.y + r.h - (placedRect.y + placedRect.h) });
                if (placedRect.x > r.x && placedRect.x < r.x + r.w) newRects.push({ x: r.x, y: r.y, w: placedRect.x - r.x, h: r.h });
                if (placedRect.x + placedRect.w < r.x + r.w) newRects.push({ x: placedRect.x + placedRect.w, y: r.y, w: r.x + r.w - (placedRect.x + placedRect.w), h: r.h });
            }
        }
        freeRects.push(...newRects);
        pruneFreeRects();
    };

    if (useSequentialPacking) {
        let i = 0;
        while (i < itemsToProcess.length) {
            const item = itemsToProcess[i];
            const wWithSpacing = item.w + spacing;
            const hWithSpacing = item.h + spacing;
            let bestScore = MAX_INT; let bestRectIndex = -1; let bestRotated = false;
            for (let j = 0; j < freeRects.length; j++) {
                const rect = freeRects[j];
                if (maxHeight > 0 && rect.y + item.h > maxHeight) continue;
                if (item.w <= rect.w && item.h <= rect.h) {
                    const score = calculateScore(rect, item.w + spacing, item.h + spacing, heuristic);
                    if (score < bestScore) { bestScore = score; bestRectIndex = j; bestRotated = false; }
                }
                if (allowRotation) {
                    if (maxHeight > 0 && rect.y + item.w > maxHeight) continue;
                    if (item.h <= rect.w && item.w <= rect.h) {
                        const score = calculateScore(rect, item.h + spacing, item.w + spacing, heuristic);
                        if (score < bestScore) { bestScore = score; bestRectIndex = j; bestRotated = true; }
                    }
                }
            }
            if (bestRectIndex !== -1) {
                placeItem(item, freeRects[bestRectIndex], bestRotated);
                itemsToProcess.splice(i, 1);
            } else i++;
        }
        unpacked.push(...itemsToProcess);
    } else {
        while (true) {
            let bestGlobalScore = MAX_INT;
            let bestMove = null;
            for (let i = 0; i < itemsToProcess.length; i++) {
                const item = itemsToProcess[i];
                for (let j = 0; j < freeRects.length; j++) {
                    const rect = freeRects[j];
                    if (maxHeight === 0 || rect.y + item.h <= maxHeight) {
                        if (item.w <= rect.w && item.h <= rect.h) {
                            let score = calculateScore(rect, item.w + spacing, item.h + spacing, heuristic);
                            if (useRandom) score += Math.random() * 20;
                            if (score < bestGlobalScore) { bestGlobalScore = score; bestMove = { itemIndex: i, rectIndex: j, rotated: false }; }
                        }
                    }
                    if (allowRotation) {
                        if (maxHeight === 0 || rect.y + item.w <= maxHeight) {
                            if (item.h <= rect.w && item.w <= rect.h) {
                                let score = calculateScore(rect, item.h + spacing, item.w + spacing, heuristic);
                                if (useRandom) score += Math.random() * 20;
                                if (score < bestGlobalScore) { bestGlobalScore = score; bestMove = { itemIndex: i, rectIndex: j, rotated: true }; }
                            }
                        }
                    }
                }
            }
            if (bestMove) {
                const item = itemsToProcess[bestMove.itemIndex];
                placeItem(item, freeRects[bestMove.rectIndex], bestMove.rotated);
                itemsToProcess.splice(bestMove.itemIndex, 1);
            } else break;
        }
        unpacked.push(...itemsToProcess);
    }

    let maxY = 0;
    packed.forEach(p => { const boxH = p.rotated ? p.w : p.h; if (p.y + boxH > maxY) maxY = p.y + boxH; });
    return { page: { pageIndex, boxes: packed, w: containerWidth, h: maxY }, unpacked };
}

const calculateScore = (rect: FreeRect, itemW: number, itemH: number, heuristic: Heuristic) => {
    const widthLeft = rect.w - itemW;
    const heightLeft = rect.h - itemH;
    if (rect.h === MAX_INT) return 1000000 + rect.y;
    switch (heuristic) {
        case 'BSSF': return Math.min(widthLeft, heightLeft);
        case 'BLSF': return Math.max(widthLeft, heightLeft);
        case 'BAF': return widthLeft * heightLeft;
        default: return Math.min(widthLeft, heightLeft);
    }
};

// --- Genetic Algorithm Helpers ---
const generatePopulation = (items: any[], size: number) => {
    const population = [];
    const indices = items.map((_, i) => i);
    const p1 = [...indices].sort((a, b) => (items[b].w * items[b].h) - (items[a].w * items[a].h)); population.push(p1);
    const p2 = [...indices].sort((a, b) => Math.max(items[b].w, items[b].h) - Math.max(items[a].w, items[a].h)); population.push(p2);
    for(let i=2; i<size; i++) {
        const p = [...indices];
        for (let j = p.length - 1; j > 0; j--) { const k = Math.floor(Math.random() * (j + 1)); [p[j], p[k]] = [p[k], p[j]]; }
        population.push(p);
    }
    return population;
};

const crossoverOX1 = (p1: number[], p2: number[]) => {
    const N = p1.length; if (N === 0) return [];
    const start = Math.floor(Math.random() * N); const end = Math.floor(Math.random() * (N - start)) + start;
    const child = new Array(N).fill(-1); const p1Set = new Set();
    for(let i=start; i<=end; i++) { child[i] = p1[i]; p1Set.add(p1[i]); }
    let current = (end + 1) % N;
    for(let i=0; i<N; i++) { const idx = (end + 1 + i) % N; const val = p2[idx]; if (!p1Set.has(val)) { child[current] = val; current = (current + 1) % N; } }
    return child;
};

const mutateSequence = (indices: number[], items: AppItem[]) => {
    if (Math.random() < 0.2) {
        const method = Math.random();
        if (method < 0.5) { const i = Math.floor(Math.random() * indices.length); const j = Math.floor(Math.random() * indices.length); [indices[i], indices[j]] = [indices[j], indices[i]]; }
        else { const i = Math.floor(Math.random() * indices.length); const j = Math.floor(Math.random() * indices.length); const val = indices.splice(i, 1)[0]; indices.splice(j, 0, val); }
    }
};

const runGA = async (
    items: AppItem[], containerW: number, containerH: number, spacing: number, allowRotation: boolean, groupByOrder: boolean,
    onProgress: (progress: number, bestLayout: FullLayout, bestCount: number) => void, shouldStop: () => boolean,
    algorithm: AlgorithmType = 'MAXRECTS'
) => {
    const POP_SIZE = 24; const GENERATIONS = 40; const ELITISM = 2;
    let population = generatePopulation(items, POP_SIZE);
    let bestLayout: FullLayout | null = null; let maxItems = 0; let minHeight = Number.MAX_VALUE;

    for (let gen = 0; gen < GENERATIONS; gen++) {
        if (shouldStop()) break;
        const evaluated = population.map(indices => {
            const ordered = indices.map(i => items[i]);
            const layout = packLayout(ordered, containerW, containerH, spacing, allowRotation, false, 'CUSTOM', 'BSSF', groupByOrder, algorithm);
            const count = layout.pages.reduce((acc, p) => acc + p.boxes.length, 0);
            return { indices, layout, count, fitness: layout.totalH };
        });
        evaluated.sort((a, b) => { if (a.count !== b.count) return b.count - a.count; return a.fitness - b.fitness; });
        const bestOfGen = evaluated[0];
        if (bestOfGen.count > maxItems || (bestOfGen.count === maxItems && bestOfGen.fitness < minHeight)) {
            maxItems = bestOfGen.count; minHeight = bestOfGen.fitness; bestLayout = bestOfGen.layout;
            onProgress(((gen+1)/GENERATIONS)*100, bestLayout, maxItems);
        } else { onProgress(((gen+1)/GENERATIONS)*100, bestLayout, maxItems); }

        const nextPop = [];
        for(let i=0; i<ELITISM; i++) nextPop.push(evaluated[i].indices);
        while(nextPop.length < POP_SIZE) {
            const p1 = evaluated[Math.floor(Math.random()*evaluated.length)].indices;
            const p2 = evaluated[Math.floor(Math.random()*evaluated.length)].indices;
            const child = crossoverOX1(p1, p2); mutateSequence(child, items);
            nextPop.push(child);
        }
        population = nextPop;
        await new Promise(r => setTimeout(r, 0));
    }
};

function drawRoundedRect(ctx, x, y, width, height, radii) {
    const { tl = 0, tr = 0, br = 0, bl = 0 } = radii;
    ctx.beginPath();
    ctx.moveTo(x + tl, y);
    ctx.lineTo(x + width - tr, y);
    ctx.arcTo(x + width, y, x + width, y + tr, tr);
    ctx.lineTo(x + width, y + height - br);
    ctx.arcTo(x + width, y + height, x + width - br, y + height, br);
    ctx.lineTo(x + bl, y + height);
    ctx.arcTo(x, y + height, x, y + height - bl, bl);
    ctx.lineTo(x, y + tl);
    ctx.arcTo(x, y, x + tl, y, tl);
    ctx.closePath();
}

interface AppSettings {
    colorKeywords: string[];
    materialKeywords: string[];
    classificationRules: { keyword: string; result: string }[];
    cornerRadiusRules: { keyword: string; value: number }[];
    productBindings: { 
        keyword: string; 
        material: string; 
        materialMode: 'fixed' | 'auto' | 'code';
        color: string; 
        colorMode: 'fixed' | 'auto' | 'code' 
    }[];
    dimensionSeparators: string;
    diameterKeywords: string;
    slashPrefix: string;
    slashSeparator: string;
    enableSpacePattern: boolean;
    enableConcatPattern: boolean;
    materialCosts: { 
        material: string; 
        cost: number; 
        unit: 'sqm' | 'm';
        width?: number;
    }[];
}

const DEFAULT_SETTINGS: AppSettings = {
    colorKeywords: ['白色', '浅灰', '深灰', '米白', '黑色', '绿灰', '羌灰', '奶白', '深绿', '太空灰', '米黄', '中灰色', '奶咖', '黑红', '杏橙', '蓝黄', '蓝色', '黑灰', '浅绿', '奶油杏', '浅蓝', '纯白', '果绿'],
    materialKeywords: ['单面绒B', '单面绒革', '单面绒', '双面革', '双面单色'],
    classificationRules: [
        { keyword: '异形', result: '异形' },
        { keyword: '定制', result: '定制' },
        { keyword: '中间椭圆', result: '中间椭圆' },
        { keyword: '椭圆', result: '椭圆' },
        { keyword: '圆形', result: '圆形' },
        { keyword: '直径', result: '圆形' },
        { keyword: '圆角', result: '圆角' }
    ],
    cornerRadiusRules: [
        { keyword: '大圆角', value: 15 },
        { keyword: '圆角', value: 10 }
    ],
    productBindings: [
        { keyword: 'BFL514', material: '原色(示例)', materialMode: 'fixed', color: '原色(示例)', colorMode: 'fixed' }
    ],
    dimensionSeparators: 'x*×-—',
    diameterKeywords: '直径',
    slashPrefix: '/',
    slashSeparator: '-',
    enableSpacePattern: true,
    enableConcatPattern: true,
    materialCosts: [
        { material: '单面绒革', cost: 15.0, unit: 'sqm' },
        { material: '太空灰', cost: 12.0, unit: 'sqm' }
    ]
};

const determineClassifications = (item: { productCode: string, notes: string, isCustom: boolean, w: number, h: number, cornerRadius: CornerRadii, pathData?: string }, settings: AppSettings): string[] => {
    const tags = new Set<string>();
    const fullText = `${item.productCode} ${item.notes}`.trim().toLowerCase();
    
    // 1. 基于设置页面的规则匹配 (所有符合的都加上)
    for (const rule of settings.classificationRules) {
        if (rule.keyword && fullText.includes(rule.keyword.toLowerCase())) {
            tags.add(rule.result);
        }
    }

    // 2. 几何识别与路径判定
    if (item.pathData) tags.add('异形');
    
    // 宽度=高度 且 满圆角 -> 判定为圆形 (允许1mm误差)
    const isFullCircle = item.w > 0 && item.w === item.h && item.cornerRadius.tl != null && item.cornerRadius.tl >= (item.w / 2 - 1);
    if (isFullCircle) tags.add('圆形');
    
    // 备注包含椭圆关键字
    if (fullText.includes('椭圆') || fullText.includes('弧形')) {
        if (fullText.includes('中间椭圆')) tags.add('中间椭圆');
        else tags.add('椭圆');
    }
    
    // 3. 定制属性判断
    if (item.isCustom || fullText.includes('定制')) tags.add('定制');

    // 4. 自动圆角识别
    const hasRounding = Object.values(item.cornerRadius).some((r: any) => r != null && r > 0);
    if (hasRounding) tags.add('圆角');
    
    // 5. 兜底逻辑：如果没有任何标签，标记为常规
    if (tags.size === 0) tags.add('常规');
    
    return Array.from(tags);
};

// 后续代码中如果有依赖单一 string 的地方，我们约定取第一个或保持数组
const determineClassification = (item: any, settings: AppSettings) => {
    const tags = determineClassifications(item, settings);
    return tags.join(','); // 使用逗号分隔存储
};

const parseOrderDataLocal = (orders: { internalOrderNumber: string, productCode: string, notes: string, quantity: number }[], settings: AppSettings): { stagedItems: any[], unprocessedItems: any[] } => {
    let allStagedItems = []; let unprocessedMap = new Map(); const processedComprehensiveNotes = new Set<string>(); 
    const colorKeywords = settings.colorKeywords;
    const materialKeywords = settings.materialKeywords;
    const numCharMap: { [key: string]: number } = { '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };

    const convertToMm = (valStr: string) => {
        if (!valStr) return 0;
        const cleanValStr = valStr.replace(/[#\s]/g, '').toLowerCase();
        // 提取纯数字部分和单位部分
        const numMatch = cleanValStr.match(/(\d+\.?\d*)/);
        if (!numMatch) return 0;
        const val = parseFloat(numMatch[1]);
        if (isNaN(val)) return 0;
        
        if (cleanValStr.includes('cm')) return parseFloat((val * 10).toFixed(1));
        if (cleanValStr.includes('mm')) return val;
        
        // 兜底逻辑：如果是很大的整数(>500)通常是mm，如果是小的通常是cm(或者用户习惯cm)
        if (Number.isInteger(val) && val >= 500) return val;
        return parseFloat((val * 10).toFixed(1)); 
    };

    const determineCornerRadii = (w_mm, h_mm, text) => {
        const MAX_REASONABLE_RADIUS_MM = 500; 
        const priorityMatch = text.match(/(?:大圆角|圆角|半径)\s*(\d+\.?\d*)\s*(cm)?/i);
        if (priorityMatch) {
            const valueStr = priorityMatch[1];
            const value = parseFloat(valueStr);
            const isCmUnit = priorityMatch[2] && priorityMatch[2].toLowerCase() === 'cm';
            if (!isNaN(value)) {
                let radius = value;
                if (isCmUnit) radius = value * 10;
                if (radius <= MAX_REASONABLE_RADIUS_MM) return { tl: radius, tr: radius, br: radius, bl: radius };
            }
        }
        if (text.includes('椭圆') || text.includes('中间椭圆') || text.includes('两头是圆弧')) {
            const radius = Math.min(w_mm, h_mm) / 2;
            return { tl: radius, tr: radius, br: radius, bl: radius };
        }
        if (text.includes('圆形') || text.includes('直径')) {
            const radius = w_mm / 2;
            return { tl: radius, tr: radius, br: radius, bl: radius };
        }
        
        for (const rule of settings.cornerRadiusRules) {
            if (text.includes(rule.keyword)) {
                return { tl: rule.value, tr: rule.value, br: rule.value, bl: rule.value };
            }
        }

        return { tl: 0, tr: 0, br: 0, bl: 0 };
    };

    orders.forEach(order => {
        const { internalOrderNumber, productCode, notes, quantity: excelQuantity } = order;
        if (!internalOrderNumber) return;
        const cleanProductCode = (productCode || '').trim();
        const cleanNotes = (notes || '').replace(/N\/A/gi, '').trim();
        const isCustomOrder = cleanNotes.includes('定制');
        const isComprehensive = /[+、，,]/.test(cleanNotes);
        const uniqueNoteIdentifier = `${internalOrderNumber}-${cleanNotes}`;

        if (isComprehensive && processedComprehensiveNotes.has(uniqueNoteIdentifier)) return;

        const fullText = `${cleanProductCode} ${cleanNotes}`.trim();
        let itemsFoundData = []; let processed = false; const notesForParsing = cleanNotes;
        
        const escapeRegex = (s: string) => s.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
        const escapedSeps = settings.dimensionSeparators.split('').map(escapeRegex).join('');
        // 允许在分隔符后面紧跟 宽/高/长/深/L/W/H 等描述词
        const dimRegexStr = `(\\d+\\.?\\d*(?:cm|mm)?)\\s*[${escapedSeps}]\\s*(?:宽|高|长|深|L|W|H)?\\s*(\\d+\\.?\\d*(?:cm|mm)?)`;
        
        const escapedDiam = settings.diameterKeywords.split('|').map(k => escapeRegex(k.trim())).join('|');
        const diameterRegexStr = `(?:直径|D|直)?\\s*(\\d+\\.?\\d*(?:cm|mm)?)\\s*(?:${escapedDiam})|(?:${escapedDiam})\\s*(\\d+\\.?\\d*(?:cm|mm)?)`;

        const dimRegexObj = new RegExp(dimRegexStr, 'i');
        const dimRegexGlobalObj = new RegExp(dimRegexStr, 'ig');
        const diameterRegexObj = new RegExp(diameterRegexStr, 'i');
        const diameterRegexGlobalObj = new RegExp(diameterRegexStr, 'ig');
        
        const dimensionRegex = dimRegexObj;
        const globalQtyRegex = /共\s*([一二两三四五六七八九十\d]+)\s*张/;
        const numDimensionGroups = (cleanNotes.match(dimRegexGlobalObj) || []).concat(cleanNotes.match(diameterRegexGlobalObj) || []).length;

        let itemParts = notesForParsing.split(/[+、，,]/).map(p => p.trim()).filter(Boolean);
        const hasMultipleParts = itemParts.length > 1;

        if (hasMultipleParts) {
            for (let i = itemParts.length - 2; i >= 0; i--) {
                const currentPart = itemParts[i];
                if (dimensionRegex.test(currentPart) || diameterRegexObj.test(currentPart)) continue;
                if (numDimensionGroups > 1 && globalQtyRegex.test(currentPart)) continue;
                itemParts[i+1] = currentPart + ' ' + itemParts[i+1];
                itemParts.splice(i, 1);
            }
        }
        
        const partsToProcess = hasMultipleParts ? itemParts.filter(p => /\d/.test(p)) : itemParts;

        if (hasMultipleParts && isComprehensive && numDimensionGroups > 1) {
            for (const part of partsToProcess) {
                let w=0, h=0, dimsFoundInPart = false; let contextText = part; let dimMatch, diameterMatch;
                diameterMatch = part.match(diameterRegexObj);
                if (diameterMatch) { 
                    const dStr = diameterMatch[1] || diameterMatch[2];
                    const d = convertToMm(dStr); w = d; h = d; dimsFoundInPart = true; contextText = contextText.replace(diameterMatch[0], ''); 
                } else {
                    dimMatch = part.match(dimRegexObj);
                    if (dimMatch) { w = convertToMm(dimMatch[1]); h = convertToMm(dimMatch[2]); dimsFoundInPart = true; contextText = contextText.replace(dimMatch[0], ''); }
                }
                if(dimsFoundInPart) {
                    processed = true; let qty = 1;
                    const qtyMatch = part.match(/共?\s*([一二两三四五六七八九十]+|\d+)\s*张/);
                    if (qtyMatch) { const qtyStr = qtyMatch[1]; const parsedNum = parseInt(qtyStr, 10); qty = numCharMap[qtyStr] || (!isNaN(parsedNum) ? parsedNum : 1); } 
                    else if (!isCustomOrder) { qty = excelQuantity || 1; }
                    let color = null; for (const c of colorKeywords) { if (part.includes(c)) { color = c; break; } }
                    itemsFoundData.push({ w, h, qty, color, partText: part, contextText: contextText });
                }
            }
        }
        
        if (!processed) {
            let w = 0, h = 0, dimsFound = false; let contextText = fullText; let dimMatch, diameterMatch, spaceDimMatch, concatDimMatch, slashDimMatch;
            diameterMatch = fullText.match(diameterRegexObj);
            if(diameterMatch) { 
                const dStr = diameterMatch[1] || diameterMatch[2];
                const d = convertToMm(dStr); w = d; h = d; dimsFound = true; contextText = contextText.replace(diameterMatch[0], ''); 
            } 
            else {
                dimMatch = fullText.match(dimRegexObj);
                if (dimMatch) { w = convertToMm(dimMatch[1]); h = convertToMm(dimMatch[2]); dimsFound = true; contextText = contextText.replace(dimMatch[0], ''); } 
                else {
                    if (settings.enableSpacePattern) {
                        const spaceDimMatch = fullText.match(/(\d{3,})\s+(\d{3,})/);
                        if (spaceDimMatch) { w = convertToMm(spaceDimMatch[1]); h = convertToMm(spaceDimMatch[2]); dimsFound = true; contextText = contextText.replace(spaceDimMatch[0], ''); } 
                    }
                    
                    if (!dimsFound) {
                        const slashPref = escapeRegex(settings.slashPrefix);
                        const slashSep = escapeRegex(settings.slashSeparator);
                        const slashDimMatch = fullText.match(new RegExp(`${slashPref}(\\d+\\.?\\d+)${slashSep}(\\d+\\.?\\d+)`));
                        if (slashDimMatch) { w = convertToMm(slashDimMatch[1]); h = convertToMm(slashDimMatch[2]); dimsFound = true; contextText = contextText.replace(slashDimMatch[0], ''); } 
                    }

                    if (!dimsFound && settings.enableConcatPattern) {
                        const concatDimMatch = fullText.match(/\b(\d{5})\b|\b(\d{4})\b/);
                        if (concatDimMatch) {
                            const dimStr = concatDimMatch[0]; let w_cm = 0, h_cm = 0;
                            if (dimStr.length === 5) { w_cm = parseInt(dimStr.substring(0, 3)); h_cm = parseInt(dimStr.substring(3)); } 
                            else if (dimStr.length === 4) { w_cm = parseInt(dimStr.substring(0, 2)); h_cm = parseInt(dimStr.substring(2)); }
                            if (w_cm > 0 && h_cm > 0) { w = w_cm * 10; h = h_cm * 10; dimsFound = true; contextText = contextText.replace(dimStr, ''); }
                        }
                    }
                }
            }
            if (dimsFound) {
                let qty = 1;
                const qtyMatch = fullText.match(/共?\s*([一二两三四五六七八九十]+|\d+)\s*张/);
                if(qtyMatch) { const qtyStr = qtyMatch[1]; const parsedNum = parseInt(qtyStr, 10); qty = numCharMap[qtyStr] || (!isNaN(parsedNum) ? parsedNum : 1); } 
                else if (!isCustomOrder) { qty = excelQuantity || 1; }
                itemsFoundData.push({w, h, qty, partText: fullText, contextText: contextText});
                processed = true;
            }
        }
        
        if (processed && itemsFoundData.length > 0) {
            let baseColor = '未知颜色'; let baseMaterial = '未知材质';

            // 0. 优先匹配产品编号映射规则
            let mappedMaterial: string | null = null;
            let mappedColor: string | null = null;

            for (const binding of settings.productBindings || []) {
                const keywords = (binding.keyword || '').toUpperCase().split(/[,\n，\s]+/).map(k => k.trim()).filter(k => k);
                const isMatched = keywords.some(key => 
                    `${cleanProductCode}`.includes(key) || 
                    `${fullText}`.toUpperCase().includes(key)
                );

                if (isMatched) {
                    if (binding.materialMode === 'fixed' && binding.material) mappedMaterial = binding.material;
                    if (binding.colorMode === 'fixed' && binding.color) mappedColor = binding.color;
                    if (mappedMaterial && mappedColor) break;
                }
            }
            if (mappedMaterial) baseMaterial = mappedMaterial;
            if (mappedColor) baseColor = mappedColor;

            if (baseColor === '未知颜色') {
                for (const color of colorKeywords) { if (fullText.includes(color)) { baseColor = color; break; } }
            }
            if (baseMaterial === '未知材质') {
                for (const material of materialKeywords) { if (fullText.includes(material)) { baseMaterial = material; break; } }
            }

            itemsFoundData.forEach(itemData => {
                const cornerRadii = determineCornerRadii(itemData.w, itemData.h, itemData.partText);
                const uniqueId = `${internalOrderNumber}-${itemData.w}x${itemData.h}-${Math.random()}`;
                const tempItemForClassification = { productCode: cleanProductCode, notes: itemData.partText, isCustom: cleanNotes.includes('定制'), w: itemData.w, h: itemData.h, cornerRadius: cornerRadii, pathData: undefined };
                const newItem = { id: uniqueId, internalOrderNumber, material: baseMaterial, color: itemData.color || baseColor, w: itemData.w, h: itemData.h, qty: itemData.qty, cornerRadius: cornerRadii, isCustom: cleanNotes.includes('定制'), productCode: cleanProductCode, notes: cleanNotes, };
                allStagedItems.push({ ...newItem, classification: determineClassification(tempItemForClassification, settings), });
            });
            if (isComprehensive) processedComprehensiveNotes.add(uniqueNoteIdentifier);
        } else {
            const isIrregular = fullText.toLowerCase().includes('异形');
            let baseColor = '未知颜色', baseMaterial = '未知材质';

            // 优先从绑定规则获取
            let mappedMaterial: string | null = null;
            let mappedColor: string | null = null;

            for (const binding of settings.productBindings || []) {
                const keywords = (binding.keyword || '').toUpperCase().split(/[,\n，\s]+/).map(k => k.trim()).filter(k => k);
                const isMatched = keywords.some(key => 
                    `${cleanProductCode}`.includes(key) || 
                    `${fullText}`.toUpperCase().includes(key)
                );

                if (isMatched) {
                    if (binding.materialMode === 'fixed' && binding.material) mappedMaterial = binding.material;
                    if (binding.colorMode === 'fixed' && binding.color) mappedColor = binding.color;
                    if (mappedMaterial && mappedColor) break;
                }
            }
            if (mappedMaterial) baseMaterial = mappedMaterial;
            if (mappedColor) baseColor = mappedColor;

            if (baseColor === '未知颜色') {
                for (const color of colorKeywords) { if (fullText.includes(color)) { baseColor = color; break; } }
            }
            if (baseMaterial === '未知材质') {
                for (const material of materialKeywords) { if (fullText.includes(material)) { baseMaterial = material; break; } }
            }
            
            if (isIrregular) {
                const qtyMatch = cleanNotes.match(/([一二两三四五六七八九十\d])张/);
                const qty = qtyMatch ? (numCharMap[qtyMatch[1]] || parseInt(qtyMatch[1], 10) || 1) : 1;
                const newItem = { id: `${internalOrderNumber}-irregular-${Math.random()}`, internalOrderNumber, material: baseMaterial, color: baseColor, w: 0, h: 0, qty: qty, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, isCustom: cleanNotes.includes('定制'), productCode: cleanProductCode, notes: cleanNotes, };
                allStagedItems.push({ ...newItem, classification: determineClassification(newItem, settings) });
            } else if (!unprocessedMap.has(internalOrderNumber)) {
                unprocessedMap.set(internalOrderNumber, { 
                    internalOrderNumber, 
                    details: { productCode: cleanProductCode, notes: cleanNotes },
                    material_manual: baseMaterial,
                    color_manual: baseColor
                });
            }
        }
    });

    return { stagedItems: allStagedItems, unprocessedItems: Array.from(unprocessedMap.values()) };
};

interface ItemPreviewProps {
    item: AppItem;
    onClick?: () => void;
}

const ItemPreview = React.memo(({ item, onClick }: ItemPreviewProps) => {
    const { w, h, cornerRadius, classification, pathData } = item;
    if (pathData) {
         return (
            <div className="preview-container" onClick={onClick}>
                <svg viewBox={`0 0 ${w + 2} ${h + 2}`} preserveAspectRatio="xMidYMid meet">
                    <path d={pathData} transform={`translate(1, 1)`} fill={'rgba(108, 117, 125, 0.2)'} stroke={'rgba(108, 117, 125, 0.8)'} strokeWidth="1" vectorEffect="non-scaling-stroke" />
                </svg>
            </div>
        );
    }
    if (!w || !h) {
        return (
             <div className="preview-container" onClick={onClick} style={{display: 'flex', alignItems: 'center', justifyContent: 'center'}}>
                <div className="preview-placeholder" style={{fontWeight: classification === '异形' ? 'bold' : 'normal', color: '#6c757d'}}>
                    {classification === '异形' ? '异形' : 'N/A'}
                </div>
            </div>
        );
    }
    const viewBoxW = w + 2;
    const viewBoxH = h + 2;
    const { tl = 0, tr = 0, br = 0, bl = 0 } = cornerRadius;
    const maxRx = w / 2;
    const maxRy = h / 2;
    const rtl = Math.min(tl, maxRx, maxRy);
    const rtr = Math.min(tr, maxRx, maxRy);
    const rbr = Math.min(br, maxRx, maxRy);
    const rbl = Math.min(bl, maxRx, maxRy);
    const pathDataRect = `M ${1 + rtl},1 L ${w + 1 - rtr},1 A ${rtr},${rtr} 0 0 1 ${w + 1},${1 + rtr} L ${w + 1},${h + 1 - rbr} A ${rbr},${rbr} 0 0 1 ${w + 1 - rbr},${h + 1} L ${1 + rbl},${h + 1} A ${rbl},${rbl} 0 0 1 1,${h + 1 - rbl} L 1,${1 + rtl} A ${rtl},${rtl} 0 0 1 ${1 + rtl},1 Z`;
    const isIrregular = classification === '异形';

    return (
        <div className="preview-container" onClick={onClick}>
            <svg viewBox={`0 0 ${viewBoxW} ${viewBoxH}`} preserveAspectRatio="xMidYMid meet">
                <path d={pathDataRect} fill={isIrregular ? 'rgba(108, 117, 125, 0.2)' : 'rgba(0, 123, 255, 0.1)'} stroke={isIrregular ? 'rgba(108, 117, 125, 0.8)' : 'rgba(0, 123, 255, 0.8)'} strokeWidth="1" vectorEffect="non-scaling-stroke" />
                 {isIrregular && <text x={viewBoxW / 2} y={viewBoxH / 2} textAnchor="middle" dominantBaseline="central" fontSize={Math.min(w, h) * 0.4} fill="#6c757d" fontWeight="bold" style={{ pointerEvents: 'none' }}>异形</text>}
            </svg>
        </div>
    );
});

interface StagedItemRowProps {
    item: AppItem;
    index: number;
    rowIndex: number;
    onChange: (index: number, field: string, value: any, cornerKey?: string) => void;
    onDelete: (id: string) => void;
    onPreviewClick: () => void;
    onMouseEnter: () => void;
    onMouseLeave: () => void;
    isHighlighted: boolean;
    isDuplicate: boolean;
}

const StagedItemRow = React.memo(({ item, index, rowIndex, onChange, onDelete, onPreviewClick, onMouseEnter, onMouseLeave, isHighlighted, isDuplicate }: StagedItemRowProps) => {
    return (
        <tr className={isHighlighted ? 'highlighted' : ''} onMouseEnter={onMouseEnter} onMouseLeave={onMouseLeave}>
            <td>{rowIndex}</td>
            <td className={isDuplicate ? 'duplicate-order' : ''}><input type="text" value={item.internalOrderNumber} onChange={(e) => onChange(index, 'internalOrderNumber', e.target.value)} /></td>
            <td><ItemPreview item={item} onClick={onPreviewClick} /></td>
            <td className="read-only" title={item.productCode}>{item.productCode}</td>
            <td className="read-only notes-cell" title={item.notes}>{item.notes}</td>
            <td>
                <div style={{display: 'flex', flexWrap: 'wrap', gap: '4px'}}>
                    {(item.classification || '').split(',').map((tag, tagIndex) => (
                        <span key={`${tag}-${tagIndex}`} style={{background: '#e7f5ff', color: '#1971c2', padding: '2px 6px', borderRadius: '4px', fontSize: '11px', whiteSpace: 'nowrap'}}>
                            {tag}
                        </span>
                    ))}
                </div>
            </td>
            <td><input type="number" value={item.qty} onChange={(e) => onChange(index, 'qty', e.target.value)} /></td>
            <td><input type="text" value={item.material} onChange={(e) => onChange(index, 'material', e.target.value)} /></td>
            <td><input type="text" value={item.color} onChange={(e) => onChange(index, 'color', e.target.value)} /></td>
            <td><input type="number" value={item.w} onChange={(e) => onChange(index, 'w', e.target.value)} /></td>
            <td><input type="number" value={item.h} onChange={(e) => onChange(index, 'h', e.target.value)} /></td>
            <td><input type="number" value={item.cornerRadius.tl} onChange={(e) => onChange(index, 'cornerRadius', e.target.value)} /></td>
            <td><button onClick={() => onDelete(item.id)} className="button button-delete">删除</button></td>
        </tr>
    );
});

const MappingModal = ({ 
    isOpen, 
    onClose, 
    itemInfo, 
    settings, 
    onUpdate 
}: { 
    isOpen: boolean; 
    onClose: () => void; 
    itemInfo: { productCode: string; notes: string } | null; 
    settings: AppSettings; 
    onUpdate: (s: AppSettings) => void;
}) => {
    const [localBindings, setLocalBindings] = useState(settings.productBindings || []);

    const [newMaterialTag, setNewMaterialTag] = useState('');
    const [newColorTag, setNewColorTag] = useState('');

    useEffect(() => {
        if (isOpen) {
            setLocalBindings(settings.productBindings || []);
            setNewMaterialTag('');
            setNewColorTag('');
        }
    }, [isOpen, settings.productBindings]);

    const handleAddMaterialTag = (val: string) => {
        if (!val || !itemInfo) return;
        const exists = localBindings.find(b => b.material === val && b.materialMode === 'fixed' && b.colorMode === 'auto');
        if (exists) {
            const codes = new Set((exists.keyword || '').split(/[,\n，\s]+/).filter(k => k));
            codes.add(itemInfo.productCode);
            const newBindings = localBindings.map(b => b === exists ? { ...b, keyword: Array.from(codes).join(', ') } : b);
            setLocalBindings(newBindings);
        } else {
            setLocalBindings([...localBindings, { keyword: itemInfo.productCode, material: val, materialMode: 'fixed', color: '', colorMode: 'auto' }]);
        }
        setNewMaterialTag('');
    };

    const handleAddColorTag = (val: string) => {
        if (!val || !itemInfo) return;
        const exists = localBindings.find(b => b.color === val && b.colorMode === 'fixed' && b.materialMode === 'auto');
        if (exists) {
            const codes = new Set((exists.keyword || '').split(/[,\n，\s]+/).filter(k => k));
            codes.add(itemInfo.productCode);
            const newBindings = localBindings.map(b => b === exists ? { ...b, keyword: Array.from(codes).join(', ') } : b);
            setLocalBindings(newBindings);
        } else {
            setLocalBindings([...localBindings, { keyword: itemInfo.productCode, color: val, colorMode: 'fixed', material: '', materialMode: 'auto' }]);
        }
        setNewColorTag('');
    };

    const materialTags = Array.from(new Set(localBindings.filter(b => b.materialMode === 'fixed' && b.colorMode === 'auto').map(b => b.material))).filter(Boolean);
    const colorTags = Array.from(new Set(localBindings.filter(b => b.colorMode === 'fixed' && b.materialMode === 'auto').map(b => b.color))).filter(Boolean);

    const handleAdd = () => {
        setLocalBindings([...localBindings, { 
            keyword: itemInfo?.productCode || '', 
            material: '', 
            materialMode: 'fixed' as const, 
            color: '', 
            colorMode: 'fixed' as const 
        }]);
    };

    const handleRemove = (index: number) => {
        setLocalBindings(localBindings.filter((_, i) => i !== index));
    };

    const handleUpdate = (index: number, field: string, value: any) => {
        const current = [...localBindings];
        current[index] = { ...current[index], [field]: value };
        setLocalBindings(current);
    };

    const handleSave = () => {
        onUpdate({ ...settings, productBindings: localBindings });
        onClose();
    };

    if (!isOpen) return null;

    return (
        <div className="modal-backdrop" style={{zIndex: 2000}}>
            <div className="modal-content" style={{maxWidth: '850px', width: '90%', maxHeight: '90vh', display: 'flex', flexDirection: 'column'}}>
                <div className="modal-header">
                    <h3 style={{margin: 0, border: 'none'}}>映射规则设置</h3>
                    <button onClick={onClose} style={{background: 'none', border: 'none', fontSize: '24px', cursor: 'pointer', color: '#666'}}>&times;</button>
                </div>
                <div className="modal-body" style={{overflowY: 'auto', padding: '20px'}}>
                    {itemInfo && (
                        <div style={{background: '#f8f9fa', padding: '12px', borderRadius: '4px', marginBottom: '15px', fontSize: '0.9rem', border: '1px solid #dee2e6'}}>
                            <p style={{margin: '0 0 8px 0'}}><strong>当前订单信息：</strong></p>
                            <div style={{display: 'flex', gap: '20px'}}>
                                <span>产品编号: <code style={{color: '#d6336c', background: '#fff0f6', padding: '2px 4px', borderRadius: '2px'}}>{itemInfo.productCode}</code></span>
                                <span>备注: <span style={{color: '#666'}}>{itemInfo.notes}</span></span>
                            </div>
                            <div style={{marginTop: '15px', display: 'flex', gap: '20px', alignItems: 'flex-start'}}>
                                <div style={{flex: 1, display: 'flex', flexDirection: 'column', gap: '8px'}}>
                                    <div style={{display: 'flex', gap: '4px'}}>
                                        <input 
                                            type="text" 
                                            placeholder="输入材质标签 (如: 硅藻泥)" 
                                            className="input-base" 
                                            style={{flex: 1, padding: '4px 8px', fontSize: '12px'}}
                                            value={newMaterialTag}
                                            onChange={(e) => setNewMaterialTag(e.target.value)}
                                        />
                                        <button 
                                            className="button button-small" 
                                            onClick={() => handleAddMaterialTag(newMaterialTag)}
                                            style={{whiteSpace: 'nowrap'}}
                                        >
                                            快捷绑定材质
                                        </button>
                                    </div>
                                    <div style={{display: 'flex', flexWrap: 'wrap', gap: '5px'}}>
                                        {materialTags.map(tag => (
                                            <span 
                                                key={tag} 
                                                onClick={() => handleAddMaterialTag(tag as string)}
                                                style={{fontSize: '10px', background: '#e7f5ff', color: '#1971c2', padding: '2px 6px', borderRadius: '10px', cursor: 'pointer', border: '1px solid #d0ebff'}}
                                            >
                                                + {tag}
                                            </span>
                                        ))}
                                    </div>
                                </div>

                                <div style={{flex: 1, display: 'flex', flexDirection: 'column', gap: '8px'}}>
                                    <div style={{display: 'flex', gap: '4px'}}>
                                        <input 
                                            type="text" 
                                            placeholder="输入颜色标签" 
                                            className="input-base" 
                                            style={{flex: 1, padding: '4px 8px', fontSize: '12px', borderColor: '#2f9e44'}}
                                            value={newColorTag}
                                            onChange={(e) => setNewColorTag(e.target.value)}
                                        />
                                        <button 
                                            className="button button-small" 
                                            style={{background: '#2f9e44', color: '#fff', whiteSpace: 'nowrap'}} 
                                            onClick={() => handleAddColorTag(newColorTag)}
                                        >
                                            快捷绑定颜色
                                        </button>
                                    </div>
                                    <div style={{display: 'flex', flexWrap: 'wrap', gap: '5px'}}>
                                        {colorTags.map(tag => (
                                            <span 
                                                key={tag} 
                                                onClick={() => handleAddColorTag(tag as string)}
                                                style={{fontSize: '10px', background: '#ebfbee', color: '#2f9e44', padding: '2px 6px', borderRadius: '10px', cursor: 'pointer', border: '1px solid #d3f9d8'}}
                                            >
                                                + {tag}
                                            </span>
                                        ))}
                                    </div>
                                </div>
                            </div>
                        </div>
                    )}
                    
                    <div style={{display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px'}}>
                        <div>
                            <h4 style={{margin: '0 0 10px 0', color: '#1971c2'}}>材质映射标签</h4>
                            <div style={{display: 'flex', flexDirection: 'column', gap: '10px'}}>
                                {localBindings.filter(b => b.materialMode === 'fixed' && b.colorMode === 'auto').map((rule, idx) => (
                                    <div key={idx} style={{background: '#f1f3f5', padding: '10px', borderRadius: '4px'}}>
                                        <div style={{display: 'flex', justifyContent: 'space-between', marginBottom: '5px'}}>
                                            <strong>{rule.material}</strong>
                                            <button onClick={() => setLocalBindings(localBindings.filter(b => b !== rule))} style={{border: 'none', background: 'none', color: '#ff4d4f', cursor: 'pointer'}}>&times;</button>
                                        </div>
                                        <textarea 
                                            value={rule.keyword} 
                                            onChange={(e) => {
                                                const newBindings = [...localBindings];
                                                const ri = newBindings.indexOf(rule);
                                                newBindings[ri] = { ...newBindings[ri], keyword: e.target.value.toUpperCase() };
                                                setLocalBindings(newBindings);
                                            }}
                                            style={{width: '100%', height: '60px', fontSize: '11px', border: '1px solid #ddd'}}
                                        />
                                    </div>
                                ))}
                            </div>
                        </div>
                        <div>
                            <h4 style={{margin: '0 0 10px 0', color: '#2f9e44'}}>颜色映射标签</h4>
                            <div style={{display: 'flex', flexDirection: 'column', gap: '10px'}}>
                                {localBindings.filter(b => b.colorMode === 'fixed' && b.materialMode === 'auto').map((rule, idx) => (
                                    <div key={idx} style={{background: '#f1f3f5', padding: '10px', borderRadius: '4px'}}>
                                        <div style={{display: 'flex', justifyContent: 'space-between', marginBottom: '5px'}}>
                                            <strong>{rule.color}</strong>
                                            <button onClick={() => setLocalBindings(localBindings.filter(b => b !== rule))} style={{border: 'none', background: 'none', color: '#ff4d4f', cursor: 'pointer'}}>&times;</button>
                                        </div>
                                        <textarea 
                                            value={rule.keyword} 
                                            onChange={(e) => {
                                                const newBindings = [...localBindings];
                                                const ri = newBindings.indexOf(rule);
                                                newBindings[ri] = { ...newBindings[ri], keyword: e.target.value.toUpperCase() };
                                                setLocalBindings(newBindings);
                                            }}
                                            style={{width: '100%', height: '60px', fontSize: '11px', border: '1px solid #ddd'}}
                                        />
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                </div>
                <div className="modal-footer" style={{padding: '15px 20px', borderTop: '1px solid #eee', display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                    <button className="button button-small" onClick={handleAdd} style={{background: '#52c41a', color: 'white'}}>新增绑定规则</button>
                    <div style={{display: 'flex', gap: '10px'}}>
                        <button className="button" onClick={onClose} style={{background: '#f0f0f0', color: '#555'}}>取消</button>
                        <button className="button button-primary" onClick={handleSave}>保存映射关系并直接同步</button>
                    </div>
                </div>
            </div>
        </div>
    );
};

const generateArcPoints = (cx: number, cy: number, radius: number, startAngle: number, endAngle: number, segments: number): {x: number, y: number}[] => {
    const points = [];
    if (segments <= 0) segments = 1;
    const angleStep = (endAngle - startAngle) / segments;
    for (let i = 0; i <= segments; i++) {
        const angle = startAngle + i * angleStep;
        const x = cx + radius * Math.cos(angle * Math.PI / 180);
        const y = cy + radius * Math.sin(angle * Math.PI / 180);
        points.push({ x, y });
    }
    return points;
};

const parsePlt = (pltContent: string) => {
    const commands = pltContent.split(/[;\r\n]+/).map(cmd => cmd.trim()).filter(Boolean);
    const strokes: {x: number, y: number}[][] = [];
    let currentStroke: {x: number, y: number}[] = [];
    let isPenDown = false;
    let currentX = 0;
    let currentY = 0;
    const addPoint = () => { if (isPenDown) currentStroke.push({ x: currentX / PLOTTER_UNITS_PER_MM, y: currentY / PLOTTER_UNITS_PER_MM }); };
    const penUp = () => { if (currentStroke.length > 1) strokes.push(currentStroke); currentStroke = []; isPenDown = false; };
    const penDown = () => { if (isPenDown && currentStroke.length > 1) strokes.push(currentStroke); isPenDown = true; currentStroke = [{ x: currentX / PLOTTER_UNITS_PER_MM, y: currentY / PLOTTER_UNITS_PER_MM }]; };

    for (const cmd of commands) {
        if (cmd.length < 2) continue;
        const instruction = cmd.substring(0, 2).toUpperCase();
        const argsStr = cmd.substring(2);
        const args = (argsStr.match(/-?\d+(\.\d+)?/g) || []).map(Number);
        switch (instruction) {
            case 'PD': penDown(); if (args.length >= 2) { for (let i = 0; i < args.length; i += 2) { if (args[i+1] === undefined) break; currentX = args[i]; currentY = args[i+1]; addPoint(); } } break;
            case 'PU': penUp(); if (args.length >= 2) { currentX = args[args.length - 2]; currentY = args[args.length - 1]; } break;
            case 'PA': if (args.length >= 2) { for (let i = 0; i < args.length; i += 2) { if (args[i+1] === undefined) break; currentX = args[i]; currentY = args[i+1]; addPoint(); } } break;
            case 'PR': if (args.length >= 2) { for (let i = 0; i < args.length; i += 2) { if (args[i+1] === undefined) break; currentX += args[i]; currentY += args[i+1]; addPoint(); } } break;
            case 'CI': 
                if (isPenDown && args.length >= 1) {
                    const radius = args[0]; const cx = currentX; const cy = currentY;
                    const segments = Math.max(72, Math.ceil(radius / PLOTTER_UNITS_PER_MM / 2));
                    const circlePoints = generateArcPoints(cx, cy, radius, 0, 360, segments);
                    strokes.push(circlePoints.map(p => ({ x: p.x / PLOTTER_UNITS_PER_MM, y: p.y / PLOTTER_UNITS_PER_MM }))); 
                } break;
            case 'AA': 
                if (isPenDown && args.length >= 3) {
                    const cx = args[0]; const cy = args[1]; const sweepAngle = args[2];
                    if (Math.abs(sweepAngle) > 0) {
                        const dx = currentX - cx; const dy = currentY - cy; const radius = Math.sqrt(dx * dx + dy * dy);
                        if (radius > 0) {
                             const startAngle = Math.atan2(dy, dx) * (180 / Math.PI); const endAngle = startAngle + sweepAngle;
                            const segments = Math.max(18, Math.ceil(Math.abs(sweepAngle) / 5));
                            const arcPoints = generateArcPoints(cx, cy, radius, startAngle, endAngle, segments);
                            if (arcPoints.length > 1) {
                                for(let i = 1; i < arcPoints.length; i++) currentStroke.push({ x: arcPoints[i].x / PLOTTER_UNITS_PER_MM, y: arcPoints[i].y / PLOTTER_UNITS_PER_MM });
                                const lastPoint = arcPoints[arcPoints.length - 1]; currentX = lastPoint.x; currentY = lastPoint.y;
                            }
                        }
                    }
                } break;
        }
    }
    penUp(); 
    if (strokes.length === 0) return null;
    const allPoints = strokes.flat();
    if (allPoints.length < 2) return null;
    const minX = Math.min(...allPoints.map(p => p.x));
    const maxX = Math.max(...allPoints.map(p => p.x));
    const minY = Math.min(...allPoints.map(p => p.y));
    const maxY = Math.max(...allPoints.map(p => p.y));
    let pathData = '';
    strokes.forEach(stroke => {
        if (stroke.length > 1) {
            const start = stroke[0]; let pathDataSegment = `M${start.x - minX} ${start.y - minY} `;
            for (let i = 1; i < stroke.length; i++) pathDataSegment += `L${stroke[i].x - minX} ${stroke[i].y - minY} `;
            const end = stroke[stroke.length - 1];
            if (Math.abs(start.x - end.x) < 0.1 && Math.abs(start.y - end.y) < 0.1) pathDataSegment += 'Z';
            pathData += pathDataSegment + ' ';
        }
    });
    return { w: Math.round(maxX - minX), h: Math.round(maxY - minY), pathData: pathData.trim(), cornerRadius: { tl: 0, tr: 0, bl: 0, br: 0 } };
};

const generatePltPathForRect = (w: number, h: number, cornerRadius: CornerRadii, offsetX: number, offsetY: number, totalHeight?: number, rotationAngle: number = 0) => {
    const maxRx = w / 2; const maxRy = h / 2;
    const { tl = 0, tr = 0, br = 0, bl = 0 } = cornerRadius;
    const r_tl = Math.min(tl, maxRx, maxRy); const r_tr = Math.min(tr, maxRx, maxRy);
    const r_br = Math.min(br, maxRx, maxRy); const r_bl = Math.min(bl, maxRx, maxRy);

    const targetSegmentLength = 0.5 * PLOTTER_UNITS_PER_MM;
    const calculateSegments = (radiusInPlotterUnits: number) => {
        if (radiusInPlotterUnits <= 0) return 0;
        return Math.min(100, Math.max(15, Math.ceil((Math.PI * radiusInPlotterUnits) / 2 / targetSegmentLength)));
    };

    const points: {x:number, y:number}[] = [];
    const pltW = w * PLOTTER_UNITS_PER_MM;
    const pltH = h * PLOTTER_UNITS_PER_MM;
    const pltTL = r_tl * PLOTTER_UNITS_PER_MM;
    const pltTR = r_tr * PLOTTER_UNITS_PER_MM;
    const pltBR = r_br * PLOTTER_UNITS_PER_MM;
    const pltBL = r_bl * PLOTTER_UNITS_PER_MM;

    const arc = (cx, cy, r, start, end) => generateArcPoints(cx, cy, r, start, end, calculateSegments(r));
    
    points.push({ x: pltTL, y: 0 });
    points.push({ x: pltW - pltTR, y: 0 });
    if (pltTR > 0) points.push(...arc(pltW - pltTR, pltTR, pltTR, 270, 360));
    points.push({ x: pltW, y: pltH - pltBR });
    if (pltBR > 0) points.push(...arc(pltW - pltBR, pltH - pltBR, pltBR, 0, 90));
    points.push({ x: pltBL, y: pltH });
    if (pltBL > 0) points.push(...arc(pltBL, pltH - pltBL, pltBL, 90, 180));
    points.push({ x: 0, y: pltTL });
    if (pltTL > 0) points.push(...arc(pltTL, pltTL, pltTL, 180, 270));
    points.push(points[0]);

    const rad = rotationAngle * Math.PI / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const cx = pltW / 2;
    const cy = pltH / 2;
    let layoutW = pltW;
    let layoutH = pltH;
    if (rotationAngle % 180 !== 0) { layoutW = pltH; layoutH = pltW; }
    const lcx = layoutW / 2;
    const lcy = layoutH / 2;

    const transformedPoints = points.map(p => {
        const ox = p.x - cx;
        const oy = p.y - cy;
        const rx = ox * cos - oy * sin;
        const ry = ox * sin + oy * cos;
        return { x: rx + lcx, y: ry + lcy };
    });

    const absX = Math.round(offsetX * PLOTTER_UNITS_PER_MM);
    const hRef = totalHeight !== undefined ? Math.round(totalHeight * PLOTTER_UNITS_PER_MM) : 0;
    
    const finalPoints = transformedPoints.map(p => {
        const px = Math.round(absX + p.x);
        let py = Math.round(offsetY * PLOTTER_UNITS_PER_MM + p.y);
        if (totalHeight !== undefined) py = hRef - py; 
        return { x: px, y: py };
    });

    if (finalPoints.length < 2) return 'PU;\n';

    let path = `PU${finalPoints[0].x},${finalPoints[0].y};\n`; // Pen Up move to start
    const BATCH_SIZE = 20; 
    for (let i = 1; i < finalPoints.length; i += BATCH_SIZE) {
        const batch = finalPoints.slice(i, i + BATCH_SIZE);
        const coords = batch.map(p => `${p.x},${p.y}`).join(',');
        path += `PD${coords};\n`; // Pen Down draw to points
    }
    path += `PU;\n`;
    return path;
};

const rotatePathDataString = (pathData: string, rotationAngle: number, originalW: number, originalH: number): string => {
    const commands = pathData.match(/([a-zA-Z])|(-?\d+(\.\d+)?)/g);
    if (!commands) return '';
    let result = ''; let idx = 0;
    const rad = rotationAngle * Math.PI / 180;
    const cos = Math.cos(rad); const sin = Math.sin(rad);
    const cx = originalW / 2; const cy = originalH / 2;
    let layoutW = originalW; let layoutH = originalH;
    if (Math.abs(rotationAngle) % 180 !== 0) { layoutW = originalH; layoutH = originalW; }
    const lcx = layoutW / 2; const lcy = layoutH / 2;

    while (idx < commands.length) {
        const token = commands[idx];
        if (/[a-zA-Z]/.test(token)) {
            result += token; idx++;
            if (token.toUpperCase() === 'Z') { result += ' '; continue; }
        } else {
            const xVal = parseFloat(commands[idx]); const yVal = parseFloat(commands[idx+1]);
            if (!isNaN(xVal) && !isNaN(yVal)) {
                const ox = xVal - cx; const oy = yVal - cy;
                const rx = ox * cos - oy * sin; const ry = ox * sin + oy * cos;
                const fx = rx + lcx; const fy = ry + lcy;
                result += `${parseFloat(fx.toFixed(3))} ${parseFloat(fy.toFixed(3))} `;
                idx += 2;
            } else idx++;
        }
    }
    return result.trim();
};

const pathDataToPlt = (pathData: string, offsetX: number, offsetY: number, totalHeight?: number, rotationAngle: number = 0, originalW: number = 0, originalH: number = 0) => {
    const rotatedPath = rotationAngle !== 0 ? rotatePathDataString(pathData, rotationAngle, originalW, originalH) : pathData;
    const commands = rotatedPath.match(/([a-zA-Z])|(-?\d+(\.\d+)?)/g);
    if (!commands) return '';
    let plt = ''; let currentCommand = '';
    const offX = Math.round(offsetX * PLOTTER_UNITS_PER_MM);
    const offY = Math.round(offsetY * PLOTTER_UNITS_PER_MM);
    const heightRef = totalHeight !== undefined ? Math.round(totalHeight * PLOTTER_UNITS_PER_MM) : 0;

    let idx = 0;
    while (idx < commands.length) {
        const token = commands[idx];
        if (/[a-zA-Z]/.test(token)) { currentCommand = token.toUpperCase(); idx++; }
        if (currentCommand === 'Z') { currentCommand = ''; continue; }
        if (currentCommand === 'M' || currentCommand === 'L') {
            const xStr = commands[idx]; const yStr = commands[idx+1];
            if (xStr && yStr && !isNaN(parseFloat(xStr)) && !isNaN(parseFloat(yStr))) {
                let px = Math.round(parseFloat(xStr) * PLOTTER_UNITS_PER_MM) + offX;
                let py = Math.round(parseFloat(yStr) * PLOTTER_UNITS_PER_MM) + offY;
                if (totalHeight !== undefined) py = heightRef - py;
                if (currentCommand === 'M') { plt += `PU${px},${py};\n`; currentCommand = 'L'; } else { plt += `PD${px},${py};\n`; }
                idx += 2;
            } else idx++;
        } else idx++;
    }
    plt += 'PU;\n';
    return plt;
};

const SettingsPage = ({ localSettings, setLocalSettings, onUpdate, onBack }: { localSettings: AppSettings, setLocalSettings: (s: AppSettings) => void, onUpdate: (s: AppSettings) => void, onBack: () => void }) => {
    const [pendingImport, setPendingImport] = useState<AppSettings | null>(null);
    const handleSave = () => {
        onUpdate(localSettings);
        onBack();
    };

    const handleAddKeyword = (key: keyof AppSettings, value: string) => {
        if (!value) return;
        const current = localSettings[key] as string[];
        if (current.includes(value)) return;
        const newSettings = { ...localSettings, [key]: [...current, value] };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleRemoveKeyword = (key: keyof AppSettings, value: string) => {
        const current = localSettings[key] as string[];
        const newSettings = { ...localSettings, [key]: current.filter(v => v !== value) };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleAddRule = (key: 'classificationRules' | 'cornerRadiusRules', rule: any) => {
        const current = localSettings[key] as any[];
        const newSettings = { ...localSettings, [key]: [...current, rule] };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleAddProductBinding = () => {
        const current = localSettings.productBindings || [];
        const newSettings = { 
            ...localSettings, 
            productBindings: [...current, { 
                keyword: '', 
                material: '', 
                materialMode: 'fixed' as const, 
                color: '', 
                colorMode: 'fixed' as const 
            }] 
        };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleRemoveProductBinding = (index: number) => {
        const current = localSettings.productBindings || [];
        const newSettings = { ...localSettings, productBindings: current.filter((_, i) => i !== index) };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const [newMaterialTagName, setNewMaterialTagName] = useState('');
    const [newColorTagName, setNewColorTagName] = useState('');

    const handleAddMaterialTagFromSettings = () => {
        if (!newMaterialTagName) return;
        const newBindings = [...(localSettings.productBindings || []), { 
            keyword: '', material: newMaterialTagName, materialMode: 'fixed' as const, color: '', colorMode: 'auto' as const 
        }];
        handleCustomUpdate('productBindings', newBindings);
        setNewMaterialTagName('');
    };

    const handleAddColorTagFromSettings = () => {
        if (!newColorTagName) return;
        const newBindings = [...(localSettings.productBindings || []), { 
            keyword: '', color: newColorTagName, colorMode: 'fixed' as const, material: '', materialMode: 'auto' as const 
        }];
        handleCustomUpdate('productBindings', newBindings);
        setNewColorTagName('');
    };

    const handleUpdateProductBinding = (index: number, field: string, value: any) => {
        const current = [...(localSettings.productBindings || [])];
        current[index] = { ...current[index], [field]: value };
        const newSettings = { ...localSettings, productBindings: current };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleRemoveRule = (key: 'classificationRules' | 'cornerRadiusRules', index: number) => {
        const current = localSettings[key] as any[];
        const newSettings = { ...localSettings, [key]: current.filter((_, i) => i !== index) };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleUpdateRule = (key: 'classificationRules' | 'cornerRadiusRules', index: number, field: string, value: any) => {
        const current = [...(localSettings[key] as any[])];
        current[index] = { ...current[index], [field]: value };
        const newSettings = { ...localSettings, [key]: current };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleUpdateMaterialCost = (index: number, field: string, value: any) => {
        const current = [...(localSettings.materialCosts || [])];
        current[index] = { ...current[index], [field]: value };
        const newSettings = { ...localSettings, materialCosts: current };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleAddMaterialCost = () => {
        const current = localSettings.materialCosts || [];
        const newSettings = { ...localSettings, materialCosts: [...current, { material: '', cost: 0, unit: 'sqm' }] };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleRemoveMaterialCost = (index: number) => {
        const current = localSettings.materialCosts || [];
        const newSettings = { ...localSettings, materialCosts: current.filter((_, i) => i !== index) };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleUpdateSimple = (key: keyof AppSettings, value: any) => {
        const newSettings = { ...localSettings, [key]: value };
        setLocalSettings(newSettings);
        onUpdate(newSettings);
    };

    const handleExport = () => {
        const blob = new Blob([JSON.stringify(localSettings, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `nesting-settings-${new Date().toISOString().split('T')[0]}.json`;
        a.click();
    };

    const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const imported = JSON.parse(event.target?.result as string);
                setPendingImport(imported);
            } catch (err) {
                alert('解析文件失败: ' + err);
            }
        };
        reader.readAsText(file);
    };

    const applyImport = (mode: 'merge' | 'overwrite') => {
        if (!pendingImport) return;

        if (mode === 'overwrite') {
            setLocalSettings(pendingImport);
            onUpdate(pendingImport);
        } else {
            const mergeArray = <T,>(existing: T[], incoming: T[], keyProp?: keyof T): T[] => {
                const res = [...existing];
                const incomingArr = Array.isArray(incoming) ? incoming : [];
                incomingArr.forEach(item => {
                    if (keyProp) {
                        const exists = res.find(r => (r as any)[keyProp] === (item as any)[keyProp]);
                        if (!exists) res.push(item);
                    } else {
                        if (!res.includes(item)) res.push(item);
                    }
                });
                return res;
            };

            const merged: AppSettings = {
                ...localSettings,
                ...pendingImport,
                colorKeywords: mergeArray(localSettings.colorKeywords, pendingImport.colorKeywords || []),
                materialKeywords: mergeArray(localSettings.materialKeywords, pendingImport.materialKeywords || []),
                classificationRules: mergeArray(localSettings.classificationRules, pendingImport.classificationRules || [], 'keyword'),
                cornerRadiusRules: mergeArray(localSettings.cornerRadiusRules, pendingImport.cornerRadiusRules || [], 'keyword'),
                productBindings: mergeArray(localSettings.productBindings || [], pendingImport.productBindings || [], 'keyword'),
            };
            setLocalSettings(merged);
            onUpdate(merged);
        }
        setPendingImport(null);
        alert('导入成功!');
    };

    return (
        <div className="container settings-container">
            {pendingImport && (
                <div className="modal-backdrop" style={{zIndex: 3000}}>
                    <div className="modal-content" style={{maxWidth: '400px', textAlign: 'center', padding: '30px'}}>
                        <h3 style={{marginBottom: '20px', border: 'none'}}>导入方式确认</h3>
                        <p style={{color: '#666', fontSize: '14px', marginBottom: '30px'}}>检测到已加载配置，请选择如何应用：</p>
                        <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                            <button className="button button-primary" onClick={() => applyImport('merge')}>合并导入 (保留现有，仅新增内容)</button>
                            <button className="button" style={{background: '#ff4d4f', color: '#fff'}} onClick={() => applyImport('overwrite')}>覆盖导入 (清除当前内容，完全替换)</button>
                            <button className="button" onClick={() => setPendingImport(null)} style={{marginTop: '10px'}}>取消</button>
                        </div>
                    </div>
                </div>
            )}
            <div className="settings-header-actions">
                 <input type="file" id="import-settings" style={{display: 'none'}} onChange={handleImport} accept=".json"/>
                <button className="button" onClick={() => document.getElementById('import-settings')?.click()}>导入配置</button>
                <button className="button" onClick={handleExport}>导出配置</button>
                <button className="button button-primary" onClick={handleSave}>保存更改</button>
                <button className="button" onClick={onBack}>返回</button>
            </div>
            
            <div className="settings-grid">
                {/* Colors */}
                <div className="settings-card">
                    <h3>颜色关键词 (Colors)</h3>
                    <div className="tag-input-container">
                        {localSettings.colorKeywords.map(tag => (
                            <span key={tag} className="tag">{tag} <span className="tag-remove" onClick={() => handleRemoveKeyword('colorKeywords', tag)}>&times;</span></span>
                        ))}
                    </div>
                    <div className="rule-item" style={{marginTop: '0.5rem'}}>
                        <input type="text" placeholder="新增颜色..." onKeyDown={(e) => { if(e.key === 'Enter') { handleAddKeyword('colorKeywords', e.currentTarget.value); e.currentTarget.value = ''; } }} />
                    </div>
                </div>

                {/* Materials */}
                <div className="settings-card">
                    <h3>材质关键词 (Materials)</h3>
                    <div className="tag-input-container">
                        {localSettings.materialKeywords.map(tag => (
                            <span key={tag} className="tag">{tag} <span className="tag-remove" onClick={() => handleRemoveKeyword('materialKeywords', tag)}>&times;</span></span>
                        ))}
                    </div>
                    <div className="rule-item" style={{marginTop: '0.5rem'}}>
                        <input type="text" placeholder="新增材质..." onKeyDown={(e) => { if(e.key === 'Enter') { handleAddKeyword('materialKeywords', e.currentTarget.value); e.currentTarget.value = ''; } }} />
                    </div>
                </div>

                {/* Classification Rules */}
                <div className="settings-card">
                    <h3>分类规则 (Classification)</h3>
                    <div className="rule-list">
                        {localSettings.classificationRules.map((rule, idx) => (
                            <div key={idx} className="rule-item">
                                <input type="text" value={rule.keyword} onChange={(e) => handleUpdateRule('classificationRules', idx, 'keyword', e.target.value)} placeholder="匹配词"/>
                                <input type="text" value={rule.result} onChange={(e) => handleUpdateRule('classificationRules', idx, 'result', e.target.value)} placeholder="结果分类"/>
                                <button className="button-icon button-icon-remove" onClick={() => handleRemoveRule('classificationRules', idx)}>&times;</button>
                            </div>
                        ))}
                        <button className="button button-small" onClick={() => handleAddRule('classificationRules', { keyword: '', result: '' })}>新增规则</button>
                    </div>
                </div>

                {/* Corner Radius Rules */}
                <div className="settings-card">
                    <h3>自动圆角规则 (Corner Radius)</h3>
                    <div className="rule-list">
                        {localSettings.cornerRadiusRules.map((rule, idx) => (
                            <div key={idx} className="rule-item">
                                <input type="text" value={rule.keyword} onChange={(e) => handleUpdateRule('cornerRadiusRules', idx, 'keyword', e.target.value)} placeholder="匹配词"/>
                                <input type="number" value={rule.value} onChange={(e) => handleUpdateRule('cornerRadiusRules', idx, 'value', parseInt(e.target.value))} placeholder="半径"/>
                                <button className="button-icon button-icon-remove" onClick={() => handleRemoveRule('cornerRadiusRules', idx)}>&times;</button>
                            </div>
                        ))}
                        <button className="button button-small" onClick={() => handleAddRule('cornerRadiusRules', { keyword: '', value: 0 })}>新增规则</button>
                    </div>
                </div>

                {/* Material Costs */}
                <div className="settings-card">
                    <h3>材质单价配置 (Material Costs)</h3>
                    <p style={{fontSize: '0.75rem', color: '#666', marginBottom: '0.5rem'}}>设置成本价。平方按面积计，米价根据宽幅自动计算长/宽米价。</p>
                    <div className="rule-list">
                        {(localSettings.materialCosts || []).map((mc, idx) => (
                            <div key={idx} className="rule-item" style={{display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center'}}>
                                <input type="text" value={mc.material} style={{flex: 2, minWidth: '120px'}} onChange={(e) => handleUpdateMaterialCost(idx, 'material', e.target.value)} placeholder="材质名称"/>
                                <select value={mc.unit || 'sqm'} onChange={(e) => handleUpdateMaterialCost(idx, 'unit', e.target.value)} style={{flex: 1, minWidth: '80px', height: '32px', padding: '0 4px', border: '1px solid #ced4da', borderRadius: '4px'}}>
                                    <option value="sqm">按平方</option>
                                    <option value="m">按米</option>
                                </select>
                                <input type="number" value={mc.cost} style={{flex: 1, minWidth: '80px'}} onChange={(e) => handleUpdateMaterialCost(idx, 'cost', parseFloat(e.target.value))} placeholder={mc.unit === 'm' ? "单价/米" : "单价/m²"} step="0.1"/>
                                {mc.unit === 'm' && (
                                    <input type="number" value={mc.width || 0} style={{flex: 1, minWidth: '100px'}} onChange={(e) => handleUpdateMaterialCost(idx, 'width', parseFloat(e.target.value))} placeholder="宽幅(mm)"/>
                                )}
                                <button className="button-icon button-icon-remove" onClick={() => handleRemoveMaterialCost(idx)}>&times;</button>
                            </div>
                        ))}
                        <button className="button button-small" onClick={handleAddMaterialCost}>新增单价配置</button>
                    </div>
                </div>

                {/* Product Bindings */}
                <div className="settings-card" style={{gridColumn: '1 / -1'}}>
                    <h3 className="section-title">产品编号映射规则 (Product Mappings)</h3>
                    <p style={{fontSize: '0.85rem', color: '#666', marginBottom: '1rem'}}>
                        作用：通过“标签”方式批量管理产品编号。设置材质标签后，匹配的订单将固定材质，颜色仍从关键词自动读取；反之亦然。支持批量粘贴（逗号/空格/换行分隔）。
                    </p>

                    <div style={{display: 'flex', flexDirection: 'column', gap: '20px'}}>
                        {/* 材质标签组 */}
                        <div style={{background: '#f8f9fa', padding: '15px', borderRadius: '8px', border: '1px solid #e9ecef'}}>
                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px'}}>
                                <strong style={{color: '#1971c2'}}>材质映射标签 (Material Tags)</strong>
                                <div style={{display: 'flex', gap: '8px'}}>
                                    <input 
                                        type="text" 
                                        placeholder="新材质名称..." 
                                        style={{fontSize: '12px', padding: '4px 8px', border: '1px solid #ced4da', borderRadius: '4px', width: '120px'}}
                                        value={newMaterialTagName}
                                        onChange={(e) => setNewMaterialTagName(e.target.value)}
                                    />
                                    <button className="button button-small" onClick={handleAddMaterialTagFromSettings}>+ 新增材质标签</button>
                                </div>
                            </div>
                            <div style={{display: 'flex', flexWrap: 'wrap', gap: '10px'}}>
                                {(localSettings.productBindings || []).filter(b => b.materialMode === 'fixed' && b.colorMode === 'auto').length === 0 && <span style={{color: '#999', fontSize: '12px'}}>暂无材质标签</span>}
                                {(localSettings.productBindings || []).filter(b => b.materialMode === 'fixed' && b.colorMode === 'auto').map((rule, idx) => {
                                    const actualIdx = localSettings.productBindings!.indexOf(rule);
                                    return (
                                        <div key={actualIdx} style={{background: '#fff', border: '1px solid #d0ebff', borderRadius: '6px', padding: '8px 12px', display: 'flex', flexDirection: 'column', gap: '8px', minWidth: '220px', flex: '1 1 220px'}}>
                                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px'}}>
                                                <span style={{fontWeight: 'bold', color: '#1971c2'}}>材质: {rule.material}</span>
                                                <button style={{border: 'none', background: 'none', color: '#f03e3e', cursor: 'pointer', fontSize: '18px'}} onClick={() => handleRemoveProductBinding(actualIdx)}>&times;</button>
                                            </div>
                                            <textarea 
                                                value={rule.keyword} 
                                                onChange={(e) => handleUpdateProductBinding(actualIdx, 'keyword', e.target.value.toUpperCase())}
                                                placeholder="粘贴产品编号..." 
                                                style={{fontSize: '11px', width: '100%', height: '60px', padding: '6px', border: '1px solid #ced4da', borderRadius: '4px', resize: 'vertical'}}
                                            />
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                        {/* 颜色标签组 */}
                        <div style={{background: '#f8f9fa', padding: '15px', borderRadius: '8px', border: '1px solid #e9ecef'}}>
                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px'}}>
                                <strong style={{color: '#2f9e44'}}>颜色映射标签 (Color Tags)</strong>
                                <div style={{display: 'flex', gap: '8px'}}>
                                    <input 
                                        type="text" 
                                        placeholder="新颜色名称..." 
                                        style={{fontSize: '12px', padding: '4px 8px', border: '1px solid #ced4da', borderRadius: '4px', width: '120px'}}
                                        value={newColorTagName}
                                        onChange={(e) => setNewColorTagName(e.target.value)}
                                    />
                                    <button className="button button-small" style={{background: '#2f9e44'}} onClick={handleAddColorTagFromSettings}>+ 新增颜色标签</button>
                                </div>
                            </div>
                            <div style={{display: 'flex', flexWrap: 'wrap', gap: '10px'}}>
                                {(localSettings.productBindings || []).filter(b => b.colorMode === 'fixed' && b.materialMode === 'auto').length === 0 && <span style={{color: '#999', fontSize: '12px'}}>暂无颜色标签</span>}
                                {(localSettings.productBindings || []).filter(b => b.colorMode === 'fixed' && b.materialMode === 'auto').map((rule, idx) => {
                                    const actualIdx = localSettings.productBindings!.indexOf(rule);
                                    return (
                                        <div key={actualIdx} style={{background: '#fff', border: '1px solid #d3f9d8', borderRadius: '6px', padding: '8px 12px', display: 'flex', flexDirection: 'column', gap: '8px', minWidth: '220px', flex: '1 1 220px'}}>
                                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px'}}>
                                                <span style={{fontWeight: 'bold', color: '#2f9e44'}}>颜色: {rule.color}</span>
                                                <button style={{border: 'none', background: 'none', color: '#f03e3e', cursor: 'pointer', fontSize: '18px'}} onClick={() => handleRemoveProductBinding(actualIdx)}>&times;</button>
                                            </div>
                                            <textarea 
                                                value={rule.keyword} 
                                                onChange={(e) => handleUpdateProductBinding(actualIdx, 'keyword', e.target.value.toUpperCase())}
                                                placeholder="粘贴产品编号..." 
                                                style={{fontSize: '11px', width: '100%', height: '60px', padding: '6px', border: '1px solid #ced4da', borderRadius: '4px', resize: 'vertical'}}
                                            />
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                        {/* 特殊双向绑定 */}
                        <div style={{background: '#fff9db', padding: '15px', borderRadius: '8px', border: '1px solid #fff3bf'}}>
                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px'}}>
                                <strong style={{color: '#e67700', fontSize: '13px'}}>高级/双向绑定 (Advanced Bindings)</strong>
                                <button className="button button-small" style={{fontSize: '10px', background: '#fab005', color: '#000'}} onClick={handleAddProductBinding}>+ 新增高级规则</button>
                            </div>
                            <div style={{display: 'flex', flexDirection: 'column', gap: '8px'}}>
                                {(localSettings.productBindings || []).filter(b => !((b.materialMode === 'fixed' && b.colorMode === 'auto') || (b.colorMode === 'fixed' && b.materialMode === 'auto'))).map((rule, idx) => {
                                    const actualIdx = localSettings.productBindings!.indexOf(rule);
                                    return (
                                        <div key={actualIdx} className="rule-item" style={{display: 'flex', gap: '10px', alignItems: 'center', padding: '8px', background: '#fff', border: '1px solid #ffe066', borderRadius: '4px'}}>
                                            <input style={{flex: 1, fontSize: '12px'}} type="text" value={rule.keyword} onChange={(e) => handleUpdateProductBinding(actualIdx, 'keyword', e.target.value.toUpperCase())} placeholder="编号..."/>
                                            <span style={{fontSize: '12px'}}>{"->"}</span>
                                            <div style={{display: 'flex', gap: '4px', alignItems: 'center'}}>
                                                <span style={{fontSize: '10px', color: '#999'}}>材质:</span>
                                                <input style={{width: '90px', fontSize: '12px'}} type="text" value={rule.material} onChange={(e) => handleUpdateProductBinding(actualIdx, 'material', e.target.value)} placeholder="材质"/>
                                            </div>
                                            <div style={{display: 'flex', gap: '4px', alignItems: 'center'}}>
                                                <span style={{fontSize: '10px', color: '#999'}}>颜色:</span>
                                                <input style={{width: '90px', fontSize: '12px'}} type="text" value={rule.color} onChange={(e) => handleUpdateProductBinding(actualIdx, 'color', e.target.value)} placeholder="颜色"/>
                                            </div>
                                            <button style={{color: '#f03e3e', border: 'none', background: 'none', cursor: 'pointer', fontSize: '16px'}} onClick={() => handleRemoveProductBinding(actualIdx)}>&times;</button>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                    </div>
                </div>

                {/* Simple Settings Patterns */}
                <div className="settings-card" style={{gridColumn: '1 / -1'}}>
                    <h3>解析识别设置 (简单模式)</h3>
                    <p style={{fontSize: '0.85rem', color: '#666', marginBottom: '1rem'}}>
                        <b>作用：</b> 系统会自动根据这些字符在备注里寻找尺寸。多个直径词请用 <code>|</code> 隔开。
                    </p>
                    
                    <div style={{display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '1.25rem'}}>
                        <div className="form-group">
                            <label>1. 尺寸连接符 (乘法格式)</label>
                            <input type="text" value={localSettings.dimensionSeparators} onChange={e => handleUpdateSimple('dimensionSeparators', e.target.value)} />
                            <span style={{fontSize: '0.75rem', color: '#999'}}>示例：<code>x*×-—</code></span>
                            <div style={{marginTop: '4px', fontSize: '10px', color: '#1971c2', background: '#e7f5ff', padding: '4px', borderRadius: '2px'}}>
                                <b>当前识别正则:</b> <code>{`(\\d+\\.?\\d*(?:cm|mm)?)\\s*[${localSettings.dimensionSeparators.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')}]\\s*(?:宽|高|长|深|L|W|H)?\\s*(\\d+\\.?\\d*(?:cm|mm)?)`}</code>
                            </div>
                        </div>
                        <div className="form-group">
                            <label>2. 直径关键词</label>
                            <input type="text" value={localSettings.diameterKeywords} onChange={e => handleUpdateSimple('diameterKeywords', e.target.value)} />
                            <span style={{fontSize: '0.75rem', color: '#999'}}>示例：<code>直径|圆形|R</code> (半角 | 隔开)</span>
                            <div style={{marginTop: '4px', fontSize: '10px', color: '#1971c2', background: '#e7f5ff', padding: '4px', borderRadius: '2px'}}>
                                <b>当前识别正则:</b> <code>{`(?:直径|D|直)?\\s*(\\d+\\.?\\d*(?:cm|mm)?)\\s*(?:${localSettings.diameterKeywords.split('|').map(k => k.trim().replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|')})|(?:${localSettings.diameterKeywords.split('|').map(k => k.trim().replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|')})\\s*(\\d+\\.?\\d*(?:cm|mm)?)`}</code>
                            </div>
                        </div>
                        <div className="form-group" style={{display: 'flex', flexDirection: 'column', gap: '0.5rem'}}>
                            <label>3. 特定斜杠格式配置</label>
                            <div style={{display: 'flex', gap: '5px'}}>
                                <input style={{width: '60px'}} type="text" value={localSettings.slashPrefix} onChange={e => handleUpdateSimple('slashPrefix', e.target.value)} placeholder="前缀"/>
                                <input style={{flex: 1}} type="text" value={localSettings.slashSeparator} onChange={e => handleUpdateSimple('slashSeparator', e.target.value)} placeholder="连接符"/>
                            </div>
                            <span style={{fontSize: '0.75rem', color: '#999'}}>示例：<code>/</code> 和 <code>-</code> 匹配 <code>/100-200</code></span>
                        </div>
                        <div className="form-group" style={{display: 'flex', gap: '1.5rem', alignItems: 'center', paddingTop: '1.5rem'}}>
                            <label className="inline"><input type="checkbox" checked={localSettings.enableSpacePattern} onChange={e => handleUpdateSimple('enableSpacePattern', e.target.checked)} /> 开启空格模式</label>
                            <label className="inline"><input type="checkbox" checked={localSettings.enableConcatPattern} onChange={e => handleUpdateSimple('enableConcatPattern', e.target.checked)} /> 开启连写模式</label>
                        </div>
                    </div>

                    <div style={{marginTop: '1.5rem', padding: '1rem', background: '#fff9db', borderRadius: '4px', border: '1px solid #ffe066'}}>
                        <h4 style={{margin: '0 0 0.5rem 0', color: '#856404', fontSize: '0.9rem'}}>⚠️ 设置注意事项：</h4>
                        <ul style={{margin: 0, paddingLeft: '1.2rem', fontSize: '0.85rem', color: '#856404', lineHeight: '1.5'}}>
                            <li><b>连接符：</b> 如果有新的尺寸连接方式，直接加在“尺寸连接符”框里即可（比如加个星号 <code>*</code>）。</li>
                            <li><b>直径词：</b> 填入备注里常出现的描述圆形的词，系统就会把提取到的一个数字当成宽和高。</li>
                            <li><b>空格/连写：</b> 如果备注很乱经常误匹配大型数字，可以考虑取消勾选这两个开关。</li>
                            <li><b>备份：</b> 修改前建议“导出配置”，备份当前的材质和规则。</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    );
};

interface ReviewEditModalProps {
    item: AppItem;
    index: number;
    onChange: (index: number, field: string, value: any, cornerKey?: string) => void;
    onClose: () => void;
}

interface MultiInsertModalProps {
    items: AppItem[];
    onConfirm: (count: number, orientation: 'default' | 'horizontal' | 'vertical') => void;
    onClose: () => void;
}

const MultiInsertModal = ({ items, onConfirm, onClose }: MultiInsertModalProps) => {
    const [count, setCount] = useState(items.length);
    const [orientation, setOrientation] = useState<'default' | 'horizontal' | 'vertical'>('default');

    if (items.length === 0) return null;
    const repItem = items[0];

    return (
        <div className="modal-backdrop" onClick={onClose} style={{zIndex: 2001}}>
            <div className="modal-content" onClick={e => e.stopPropagation()} style={{maxWidth: '400px'}}>
                <div className="modal-header">
                    <h4>多件插入设置: {repItem.internalOrderNumber}</h4>
                    <button onClick={onClose} className="button-close">&times;</button>
                </div>
                <div className="modal-body">
                    <div style={{display: 'flex', gap: '1rem', marginBottom: '1.5rem', alignItems: 'center'}}>
                        <div style={{width: '60px', height: '60px', border: '1px solid #eee', padding: '5px'}}>
                            <ItemPreview item={repItem} />
                        </div>
                        <div>
                            <p style={{margin: 0, fontWeight: 'bold'}}>{repItem.internalOrderNumber}</p>
                            <p style={{margin: 0, fontSize: '0.85rem', color: '#666'}}>{repItem.w} x {repItem.h} mm</p>
                            <p style={{margin: 0, fontSize: '0.85rem', color: '#666'}}>总计: {items.length} 件</p>
                        </div>
                    </div>

                    <div className="form-group">
                        <label>插入数量 (默认全部)</label>
                        <input 
                            type="number" 
                            min="1" 
                            max={items.length} 
                            value={count} 
                            onChange={e => setCount(Math.min(items.length, Math.max(1, parseInt(e.target.value) || 1)))}
                        />
                    </div>

                    <div className="form-group">
                        <label>插入方式 (物体方向)</label>
                        <div style={{display: 'flex', gap: '0.5rem'}}>
                            <button 
                                className={`button button-small ${orientation === 'default' ? 'active' : ''}`}
                                onClick={() => setOrientation('default')}
                                style={{flex: 1, backgroundColor: orientation === 'default' ? '#1971c2' : '', color: orientation === 'default' ? 'white' : ''}}
                            >
                                默认
                            </button>
                            <button 
                                className={`button button-small ${orientation === 'horizontal' ? 'active' : ''}`}
                                onClick={() => setOrientation('horizontal')}
                                style={{flex: 1, backgroundColor: orientation === 'horizontal' ? '#2f9e44' : '', color: orientation === 'horizontal' ? 'white' : ''}}
                            >
                                强制横
                            </button>
                            <button 
                                className={`button button-small ${orientation === 'vertical' ? 'active' : ''}`}
                                onClick={() => setOrientation('vertical')}
                                style={{flex: 1, backgroundColor: orientation === 'vertical' ? '#e67e22' : '', color: orientation === 'vertical' ? 'white' : ''}}
                            >
                                强制竖
                            </button>
                        </div>
                        <p style={{fontSize: '0.75rem', color: '#999', marginTop: '0.5rem'}}>
                            提示: 选中单元格后将优先在其右侧连续排列
                        </p>
                    </div>
                </div>
                <div className="modal-footer">
                    <button onClick={onClose} className="button" style={{backgroundColor: '#6c757d', color: 'white'}}>取消</button>
                    <button onClick={() => onConfirm(count, orientation)} className="button button-primary">确认插入</button>
                </div>
            </div>
        </div>
    );
};

const ReviewEditModal = ({ item, index, onChange, onClose }: ReviewEditModalProps) => {
    if (!item) return null;
    const uploadRef = useRef(null);

    const handleDownloadPlt = () => {
        const plt = generatePltPathForRect(item.w, item.h, item.cornerRadius, 0, 0, undefined, item.rotationAngle || 0);
        const blob = new Blob([plt], { type: 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `${item.internalOrderNumber || 'item'}.plt`;
        a.click();
        URL.revokeObjectURL(a.href);
    };

    const handleFileUpload = (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result as string;
            try {
                const parsedGeo = parsePlt(content);
                if (parsedGeo) {
                    onChange(index, 'w', parsedGeo.w); onChange(index, 'h', parsedGeo.h); onChange(index, 'cornerRadius', 0); 
                    onChange(index, 'pathData', parsedGeo.pathData); onChange(index, 'classification', '异形'); 
                } else alert('PLT文件解析失败，请确保文件包含有效的HPGL指令。');
            } catch (error) { alert(`解析PLT时出错: ${error.message}`); }
        };
        reader.readAsText(file);
    };

    return (
        <div className="modal-backdrop" onClick={onClose}>
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h4>编辑物件: {item.internalOrderNumber}</h4>
                    <button onClick={onClose} className="button-close">&times;</button>
                </div>
                <div className="modal-body review-modal-body">
                    <div className="form-group-grid">
                        <div className="form-group"><label>宽度 (mm)</label><input type="number" value={item.w} onChange={e => onChange(index, 'w', e.target.value)} /></div>
                        <div className="form-group"><label>高度 (mm)</label><input type="number" value={item.h} onChange={e => onChange(index, 'h', e.target.value)} /></div>
                        <div className="form-group full-width"><label>统一圆角 (mm)</label><input type="number" value={item.cornerRadius.tl} onChange={e => onChange(index, 'cornerRadius', e.target.value)} disabled={!!item.pathData} /></div>
                        <div className="form-group"><label>左上 (TL)</label><input type="number" value={item.cornerRadius.tl} onChange={e => onChange(index, 'cornerRadius', e.target.value, 'tl')} disabled={!!item.pathData}/></div>
                         <div className="form-group"><label>右上 (TR)</label><input type="number" value={item.cornerRadius.tr} onChange={e => onChange(index, 'cornerRadius', e.target.value, 'tr')} disabled={!!item.pathData}/></div>
                         <div className="form-group"><label>左下 (BL)</label><input type="number" value={item.cornerRadius.bl} onChange={e => onChange(index, 'cornerRadius', e.target.value, 'bl')} disabled={!!item.pathData}/></div>
                         <div className="form-group"><label>右下 (BR)</label><input type="number" value={item.cornerRadius.br} onChange={e => onChange(index, 'cornerRadius', e.target.value, 'br')} disabled={!!item.pathData}/></div>
                        <div className="form-group full-width button-group">
                            <input type="file" ref={uploadRef} style={{ display: 'none' }} onChange={handleFileUpload} accept=".plt"/>
                            <button className="button button-secondary" onClick={() => uploadRef.current?.click()}>上传 PLT</button>
                            <button className="button" onClick={handleDownloadPlt} disabled={!!item.pathData}>下载 PLT</button>
                        </div>
                    </div>
                    <div className="modal-preview"><ItemPreview item={item} /></div>
                </div>
                <div className="modal-footer"><button onClick={onClose} className="button button-primary">完成</button></div>
            </div>
        </div>
    );
};

// --- New Material Box Component ---
const MaterialBox = ({ title, items, onSelect, selectedGroupKey, onDoubleClick, onClose, setMultiInsertConfig, isPreset = false }) => {
    const [position, setPosition] = useState({ x: isPreset ? 100 : window.innerWidth - 320, y: 150 });
    const [size, setSize] = useState({ w: 300, h: 400 });
    const [isDragging, setIsDragging] = useState(false);
    const [dragStart, setDragStart] = useState({ x: 0, y: 0 });

    const handleMouseDown = (e) => {
        if (e.target.closest('.resize-handle') || e.target.closest('.button-close')) return;
        setIsDragging(true);
        setDragStart({ x: e.clientX - position.x, y: e.clientY - position.y });
    };

    useEffect(() => {
        const handleMouseMove = (e) => {
            if (isDragging) {
                setPosition({ x: e.clientX - dragStart.x, y: e.clientY - dragStart.y });
            }
        };
        const handleMouseUp = () => setIsDragging(false);
        if (isDragging) {
            window.addEventListener('mousemove', handleMouseMove);
            window.addEventListener('mouseup', handleMouseUp);
        }
        return () => {
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseup', handleMouseUp);
        };
    }, [isDragging, dragStart]);

    // Group items by InternalOrderNumber + Width + Height
    const groupedItems = useMemo(() => {
        if (isPreset) {
             const groups = {};
             items.forEach((item, idx) => {
                 const key = `preset-${idx}`;
                 groups[key] = [item];
             });
             return groups;
        }
        const groups = {};
        items.forEach(item => {
            const key = `${item.internalOrderNumber}|${item.w}|${item.h}`;
            if (!groups[key]) groups[key] = [];
            groups[key].push(item);
        });
        return groups;
    }, [items, isPreset]);

    const groupKeys = Object.keys(groupedItems).sort((a, b) => {
        const itemA = groupedItems[a][0];
        const itemB = groupedItems[b][0];
        if (!itemA.internalOrderNumber || !itemB.internalOrderNumber) return 0;
        return itemA.internalOrderNumber.localeCompare(itemB.internalOrderNumber, 'zh-CN', { numeric: true });
    });

    return (
        <div style={{
            position: 'fixed', left: position.x, top: position.y, width: size.w, height: size.h,
            backgroundColor: 'rgba(255, 255, 255, 0.95)', border: '1px solid #ccc',
            boxShadow: '0 4px 15px rgba(0,0,0,0.2)', borderRadius: '8px', zIndex: 1000,
            display: 'flex', flexDirection: 'column'
        }}>
            <div onMouseDown={handleMouseDown} style={{
                padding: '10px', borderBottom: '1px solid #eee', cursor: 'grab',
                backgroundColor: '#f8f9fa', borderRadius: '8px 8px 0 0', fontWeight: 'bold', userSelect: 'none',
                display: 'flex', justifyContent: 'space-between', alignItems: 'center'
            }}>
                <span>{title} ({items.length})</span>
                <button className="button-close" onClick={onClose} style={{fontSize: '1.2rem', padding: '0 5px', border: 'none', background: 'none', cursor: 'pointer'}}>&times;</button>
            </div>
            <div style={{
                flex: 1, overflowY: 'auto', padding: '10px',
                display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(80px, 1fr))', gap: '10px', alignContent: 'start'
            }}>
                {groupKeys.map(key => {
                    const group = groupedItems[key];
                    const repItem = group[0];
                    const count = group.length;
                    const isSelected = selectedGroupKey === key;
                    const color = stringToColor(repItem.internalOrderNumber || 'preset');

                    return (
                        <div key={key} 
                             onClick={() => onSelect(key, group)}
                             onDoubleClick={() => {
                                 if (isPreset) {
                                     onDoubleClick(group);
                                 } else {
                                     setMultiInsertConfig({ open: true, items: group });
                                 }
                             }}
                             draggable="true"
                             onDragStart={(e) => {
                                 e.dataTransfer.setData("application/json", JSON.stringify({ key, isPreset }));
                             }}
                             style={{
                                 border: isSelected ? '2px solid #007bff' : '1px solid #ddd',
                                 borderRadius: '4px', padding: '5px', cursor: 'pointer',
                                 backgroundColor: isSelected ? '#e7f1ff' : 'white',
                                 display: 'flex', flexDirection: 'column', alignItems: 'center',
                                 position: 'relative'
                             }}>
                            {!isPreset && count > 1 && (
                                <div style={{
                                    position: 'absolute', top: '-8px', right: '-8px',
                                    backgroundColor: '#dc3545', color: 'white', borderRadius: '50%',
                                    width: '20px', height: '20px', fontSize: '11px',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                                    fontWeight: 'bold', boxShadow: '0 2px 4px rgba(0,0,0,0.2)',
                                    zIndex: 1
                                }}>{count}</div>
                            )}
                            <div style={{width: '100%', height: '50px', marginBottom: '5px'}}>
                                <ItemPreview item={repItem as AppItem} />
                            </div>
                            <div style={{fontSize: '10px', textAlign: 'center', width: '100%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'}}>
                                {repItem.internalOrderNumber}
                            </div>
                            <div style={{fontSize: '9px', textAlign: 'center', width: '100%', color: '#666', marginTop: '1px'}}>
                                {repItem.w}x{repItem.h}
                            </div>
                            <div style={{width: '100%', height: '4px', backgroundColor: color, marginTop: '2px', borderRadius: '2px'}}></div>
                        </div>
                    );
                })}
                {items.length === 0 && <div style={{gridColumn: '1 / -1', textAlign: 'center', color: '#999', marginTop: '20px'}}>暂无物料</div>}
            </div>
            <div className="resize-handle" style={{
                position: 'absolute', bottom: 0, right: 0, width: '15px', height: '15px', cursor: 'nwse-resize',
                background: 'linear-gradient(135deg, transparent 50%, #ccc 50%)'
            }} onMouseDown={(e) => {
                e.stopPropagation();
                const startX = e.clientX; const startY = e.clientY;
                const startW = size.w; const startH = size.h;
                const onMove = (em) => setSize({ w: Math.max(200, startW + em.clientX - startX), h: Math.max(200, startH + em.clientY - startY) });
                const onUp = () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
                window.addEventListener('mousemove', onMove); window.addEventListener('mouseup', onUp);
            }}/>
        </div>
    );
};

const App = () => {
    const [settings, setSettings] = useState<AppSettings>(() => {
        const saved = localStorage.getItem('nesting_app_settings');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                const merged = { ...DEFAULT_SETTINGS, ...parsed };
                // 增强鲁棒性：如果关键配置项为空，则从默认配置中恢复，解决用户提到的“内置规则没了”的问题
                if (!merged.classificationRules || merged.classificationRules.length === 0) merged.classificationRules = DEFAULT_SETTINGS.classificationRules;
                if (!merged.colorKeywords || merged.colorKeywords.length === 0) merged.colorKeywords = DEFAULT_SETTINGS.colorKeywords;
                if (!merged.materialKeywords || merged.materialKeywords.length === 0) merged.materialKeywords = DEFAULT_SETTINGS.materialKeywords;
                if (!merged.productBindings) {
                    merged.productBindings = DEFAULT_SETTINGS.productBindings;
                } else {
                    // 确保旧有的规则也有 Mode 字段
                    merged.productBindings = merged.productBindings.map((b: any) => ({
                        ...b,
                        materialMode: b.materialMode || 'fixed',
                        colorMode: b.colorMode || 'fixed'
                    }));
                }
                return merged;
            } catch (e) {
                console.error("Failed to parse settings", e);
                return DEFAULT_SETTINGS;
            }
        }
        return DEFAULT_SETTINGS;
    });
    const [isSettingsOpen, setIsSettingsOpen] = useState(false);
    const [localSettings, setLocalSettings] = useState<AppSettings>(settings);

    // 当全局 settings 被外部（如智能同步）更新时，同步给 localSettings
    useEffect(() => {
        setLocalSettings(settings);
    }, [settings]);

    useEffect(() => {
        localStorage.setItem('nesting_app_settings', JSON.stringify(settings));
    }, [settings]);

    const [file, setFile] = useState(null);
    const [materialWidth, setMaterialWidth] = useState(1500);
    const [materialLength, setMaterialLength] = useState(0); // 0 means infinite
    const [spacing, setSpacing] = useState(10);
    const [nestingAlgorithm, setNestingAlgorithm] = useState<AlgorithmType>('MAXRECTS');
    const [isLibOpen, setIsLibOpen] = useState(false);
    const [strategyLibrary, setStrategyLibrary] = useState<SavedStrategy[]>([]);

    const exportLibrary = () => {
        const data = JSON.stringify(strategyLibrary, null, 2);
        const blob = new Blob([data], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `nesting_strategies_${new Date().toISOString().slice(0,10)}.json`;
        a.click();
        URL.revokeObjectURL(url);
    };

    const importLibrary = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const imported = JSON.parse(event.target?.result as string);
                if (Array.isArray(imported)) {
                    setStrategyLibrary(imported);
                    saveStrategyLibrary(imported);
                    alert(`成功导入 ${imported.length} 个策略`);
                }
            } catch (error) {
                alert('导入失败: 格式错误');
            }
        };
        reader.readAsText(file);
    };

    const deleteStrategy = (id: string) => {
        const updated = strategyLibrary.filter(s => s.id !== id);
        setStrategyLibrary(updated);
        saveStrategyLibrary(updated);
    };

    const clearLibrary = () => {
        if (confirm('确定要清空所有策略记录吗？')) {
            setStrategyLibrary(defaultStrategies);
            saveStrategyLibrary(defaultStrategies);
        }
    };

    useEffect(() => {
        setStrategyLibrary(loadStrategyLibrary());
    }, []);

    const learnFromLayout = useCallback((layout: FullLayout, alg: AlgorithmType, sort: SortStrategy, heur: Heuristic, rot: boolean) => {
        // Calculate total utilization
        let totalUsedArea = 0;
        let totalSheetArea = 0;
        layout.pages.forEach(p => {
            p.boxes.forEach(b => {
                totalUsedArea += b.w * b.h;
            });
            totalSheetArea += p.w * p.h;
        });
        
        const utilization = totalSheetArea > 0 ? totalUsedArea / totalSheetArea : 0;
        
        if (utilization >= 0.90) {
            setStrategyLibrary(prev => {
                // Similarity check: avoid duplicates
                const isDuplicate = prev.some(s => s.algorithm === alg && s.sortStrategy === sort && s.heuristic === heur && s.allowRotation === rot);
                if (isDuplicate) return prev;

                const newStrategy: SavedStrategy = {
                    id: `strat-${Date.now()}`,
                    name: `高效策略-${Math.round(utilization * 100)}%`,
                    algorithm: alg,
                    sortStrategy: sort,
                    heuristic: heur,
                    allowRotation: rot,
                    utilization,
                    usageCount: 1,
                    createdAt: Date.now()
                };
                
                const updated = [...prev, newStrategy].sort((a, b) => b.utilization - a.utilization);
                // Auto-cleanup: keep top 50 strategies
                const final = updated.slice(0, 50);
                saveStrategyLibrary(final);
                return final;
            });
        }
    }, []);
    const [allowRotation, setAllowRotation] = useState(true);
    const [useRandom, setUseRandom] = useState(false);
    const [groupByOrder, setGroupByOrder] = useState(false);

    const [stagedItems, setStagedItems] = useState<AppItem[]>([]);
    const [unprocessedItems, setUnprocessedItems] = useState([]);
    const [groupedOrders, setGroupedOrders] = useState(null);
    const [orderDetails, setOrderDetails] = useState(new Map());

    const [currentStep, setCurrentStep] = useState(1);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState('');
    const [isNesting, setIsNesting] = useState(false);
    const [nestingProgress, setNestingProgress] = useState(0);
    const [selectedGroupKey, setSelectedGroupKey] = useState(null);
    const [layout, setLayout] = useState<FullLayout | null>(null);
    const [mergeOrderNumbers, setMergeOrderNumbers] = useState(true);
    const [unitCosts, setUnitCosts] = useState<Record<string, { cost: number; unit: 'sqm' | 'm'; width?: number }>>({});
    const allBoxes = useMemo(() => layout ? layout.pages.flatMap(p => p.boxes) : [], [layout]);
    const uniqueMaterials = useMemo(() => { if (!layout) return []; const mats = new Set(allBoxes.map(b => b.material)); return Array.from(mats).sort(); }, [allBoxes, layout]);
    const [isReviewFullScreen, setIsReviewFullScreen] = useState(false);
    const [editingReviewItem, setEditingReviewItem] = useState<{item: AppItem, index: number} | null>(null);
    const [activeFilter, setActiveFilter] = useState('全部');
    const [isCanvasFullScreen, setIsCanvasFullScreen] = useState(false);
    
    // Deep search stats
    const [isOptimizing, setIsOptimizing] = useState(false);
    const [optimizationProgress, setOptimizationProgress] = useState(0);
    const [statusText, setStatusText] = useState(''); // New status text for detailed progress
    const [bestStats, setBestStats] = useState({ count: 0, height: 0 });
    const stopDeepSearchRef = useRef(false);

    // Selection State
    const [selectedItem, setSelectedItem] = useState<AppItem | null>(null); // Kept for Edit Modal / single focus
    const [selectedItemIds, setSelectedItemIds] = useState<Set<string>>(new Set()); // Multiple selection
    const [editedItem, setEditedItem] = useState(null);
    const [isEditModalOpen, setIsEditModalOpen] = useState(false);
    const [highlightedItemId, setHighlightedItemId] = useState(null);
    const [hoveredItemId, setHoveredItemId] = useState(null);
    const [hoveredOrderNumber, setHoveredOrderNumber] = useState<string | null>(null);
    const [mappingModal, setMappingModal] = useState<{ open: boolean; itemIndex: number | null }>({ open: false, itemIndex: null });
    
    // Canvas Navigation & Drag
    const [canvasTransform, setCanvasTransform] = useState({ scale: 1, x: 0, y: 0 });
    const [isPanning, setIsPanning] = useState(false);
    const [panStart, setPanStart] = useState({ x: 0, y: 0 });
    const [draggingItemId, setDraggingItemId] = useState<string | null>(null);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [guideLines, setGuideLines] = useState<{x1:number, y1:number, x2:number, y2:number}[]>([]); 

    // Box Selection State
    const [isBoxSelecting, setIsBoxSelecting] = useState(false);
    const [selectionBox, setSelectionBox] = useState<{x: number, y: number, w: number, h: number} | null>(null);

    const [selectedExportColumns, setSelectedExportColumns] = useState(EXPORT_COLUMNS.map(c => c.key));

    // Sync unitCosts from settings.materialCosts on load and when settings change
    useEffect(() => {
        if (settings.materialCosts) {
            const newCosts: Record<string, { cost: number; unit: 'sqm' | 'm'; width?: number }> = {};
            settings.materialCosts.forEach(mc => {
                if (mc.material) newCosts[mc.material] = { cost: mc.cost, unit: mc.unit || 'sqm', width: mc.width };
            });
            setUnitCosts(newCosts);
        }
    }, [settings.materialCosts]);

    // Handle unit cost change from UI
    const updateMaterialCostSetting = useCallback((material: string, field: string, value: any) => {
        // Update local state for immediate feedback
        setUnitCosts(prev => {
            const current = prev[material] || { cost: 0, unit: 'sqm' };
            return { ...prev, [material]: { ...current, [field]: value } };
        });
        
        // Update settings for persistence
        setSettings(prev => {
            const currentCosts = [...(prev.materialCosts || [])];
            const idx = currentCosts.findIndex(mc => mc.material === material);
            if (idx >= 0) {
                currentCosts[idx] = { ...currentCosts[idx], [field]: value };
            } else {
                currentCosts.push({ material, cost: 0, unit: 'sqm', [field]: value } as any);
            }
            return { ...prev, materialCosts: currentCosts };
        });
    }, [setSettings]);

    const handleUnitCostChange = useCallback((material: string, val: string) => {
        const num = val === '' ? 0 : parseFloat(val);
        const safeNum = isNaN(num) ? 0 : num;
        updateMaterialCostSetting(material, 'cost', safeNum);
    }, [updateMaterialCostSetting]);

    // Auto-sync new materials found in layout to settings
    useEffect(() => {
        if (uniqueMaterials.length > 0) {
            setSettings(prev => {
                const currentCosts = [...(prev.materialCosts || [])];
                let changed = false;
                uniqueMaterials.forEach(mat => {
                    if (mat && !currentCosts.some(mc => mc.material === mat)) {
                        currentCosts.push({ material: mat, cost: 0, unit: 'sqm' });
                        changed = true;
                    }
                });
                if (changed) return { ...prev, materialCosts: currentCosts };
                return prev;
            });
        }
    }, [uniqueMaterials]);

    // Material Box State
    const [materialBoxItems, setMaterialBoxItems] = useState<AppItem[]>([]);
    const [isMaterialBoxVisible, setIsMaterialBoxVisible] = useState(false);
    const [selectedBoxGroupKey, setSelectedBoxGroupKey] = useState<string | null>(null); // Track selected group key
    const [selectedBoxItems, setSelectedBoxItems] = useState<AppItem[]>([]); // Track selected items in box
    const [layoutHistory, setLayoutHistory] = useState<{layout: FullLayout, boxItems: AppItem[]}[]>([]);

    // Preset Box State
    const [isPresetBoxVisible, setIsPresetBoxVisible] = useState(false);
    const [selectedPresetGroupKey, setSelectedPresetGroupKey] = useState<string | null>(null);
    const [selectedPresetItems, setSelectedPresetItems] = useState<AppItem[]>([]);

    const canvasRef = useRef(null);
    const canvasContainerRef = useRef(null);
    
    const resetState = () => {
        setFile(null); setError(''); setGroupedOrders(null); setStagedItems([]); setUnprocessedItems([]);
        setCurrentStep(1); setSelectedGroupKey(null); setLayout(null); setSelectedItem(null); setEditedItem(null);
        setIsEditModalOpen(false); setEditingReviewItem(null); setHighlightedItemId(null); setIsReviewFullScreen(false);
        setIsCanvasFullScreen(false); setCanvasTransform({ scale: 1, x: 0, y: 0 }); setIsOptimizing(false);
        setOptimizationProgress(0); setStatusText(''); setBestStats({ count: 0, height: 0 }); setGuideLines([]); 
        const initialCosts: Record<string, { cost: number; unit: 'sqm' | 'm'; width?: number }> = {};
        if (settings.materialCosts) {
            settings.materialCosts.forEach(mc => {
                if (mc.material) initialCosts[mc.material] = { cost: mc.cost, unit: mc.unit || 'sqm', width: mc.width };
            });
        }
        setUnitCosts(initialCosts);
        setMaterialBoxItems([]); setIsMaterialBoxVisible(false); setLayoutHistory([]);
        setIsPresetBoxVisible(false); setSelectedPresetGroupKey(null); setSelectedPresetItems([]);
        setSelectedItemIds(new Set()); setSelectionBox(null); setIsBoxSelecting(false);
        setSelectedExportColumns(EXPORT_COLUMNS.map(c => c.key));
        const fileInput = document.getElementById('file-upload') as HTMLInputElement;
        if (fileInput) fileInput.value = '';
    }

    const pushHistory = (currentLayout: FullLayout, currentBoxItems: AppItem[]) => {
        setLayoutHistory(prev => {
            const newHistory = [...prev, { layout: JSON.parse(JSON.stringify(currentLayout)), boxItems: [...currentBoxItems] }];
            if (newHistory.length > 20) newHistory.shift(); // Limit history
            return newHistory;
        });
    };

    const [multiInsertConfig, setMultiInsertConfig] = useState<{ open: boolean, items: AppItem[] }>({ open: false, items: [] });

    const handleBoxItemInsert = (itemsToInsert: AppItem | AppItem[], forceRotation: boolean | null = null, isPreset: boolean = false, countToInsert?: number, orientation?: 'default' | 'horizontal' | 'vertical') => {
        const fullItems = Array.isArray(itemsToInsert) ? itemsToInsert : [itemsToInsert];
        if (fullItems.length === 0 || !layout) return;

        const count = countToInsert !== undefined ? Math.min(countToInsert, fullItems.length) : fullItems.length;
        const items = fullItems.slice(0, count);
        
        pushHistory(layout, materialBoxItems); // Save state once

        let currentPages = [...layout.pages];
        let currentInsertionPoint: { pageIndex: number, x: number, y: number } | null = null;
        
        if (selectedItem) {
            const sourcePage = currentPages.find(p => p.boxes.some(b => b.id === selectedItem.id));
            if (sourcePage) {
                const itemW = selectedItem.rotated ? selectedItem.h : selectedItem.w;
                currentInsertionPoint = {
                    pageIndex: sourcePage.pageIndex,
                    x: selectedItem.x + itemW + spacing,
                    y: selectedItem.y
                };
            }
        }

        items.forEach(item => {
            let targetPage: PageLayout;
            let finalX = 0;
            let finalY = 0;

            // Determine rotation
            let rotated = false;
            if (orientation === 'horizontal') {
                rotated = item.h > item.w;
            } else if (orientation === 'vertical') {
                rotated = item.w > item.h;
            } else if (forceRotation !== null) {
                rotated = forceRotation;
            } else {
                const w = item.w; const h = item.h;
                if (w > materialWidth && h <= materialWidth) rotated = true;
                else if (h > materialWidth && w <= materialWidth) rotated = false;
                else if (allowRotation && h > w) rotated = true;
                else if (w > h) rotated = false; 
            }

            const itemW = rotated ? item.h : item.w;
            const itemH = rotated ? item.w : item.h;

            if (currentInsertionPoint) {
                targetPage = currentPages.find(p => p.pageIndex === currentInsertionPoint!.pageIndex) || currentPages[currentPages.length-1];
                finalX = currentInsertionPoint.x;
                finalY = currentInsertionPoint.y;

                // Check bounds: if too wide, wrap to next line or resume standard flow
                if (finalX + itemW > materialWidth) {
                    finalX = 0;
                    let maxY = 0;
                    targetPage.boxes.forEach(b => {
                        const bH = b.rotated ? b.w : b.h;
                        if (b.y + bH > maxY) maxY = b.y + bH;
                    });
                    finalY = maxY > 0 ? maxY + spacing : 0;
                    currentInsertionPoint = null; // Exit right-aligned flow if it wraps
                }
            } else {
                targetPage = currentPages[currentPages.length - 1] || { pageIndex: 1, boxes: [], w: materialWidth, h: 0 };
                if (!currentPages.length) currentPages.push(targetPage);
                
                let maxY = 0;
                targetPage.boxes.forEach(b => {
                    const bH = b.rotated ? b.w : b.h;
                    if (b.y + bH > maxY) maxY = b.y + bH;
                });
                finalY = maxY > 0 ? maxY + spacing : 0;
                finalX = 0;
            }
            
            let placed = false;
            let loopCount = 0;
            while (!placed && loopCount < 50) { 
                loopCount++;
                let maxY = 0;
                targetPage.boxes.forEach(b => {
                    const bH = b.rotated ? b.w : b.h;
                    if (b.y + bH > maxY) maxY = b.y + bH;
                });

                if (targetPage.boxes.length > 0 && finalY < maxY + spacing && !currentInsertionPoint) {
                    finalY = maxY + spacing;
                }
                
                const fits = (finalX + itemW <= materialWidth) && (materialLength === 0 || finalY + itemH <= materialLength);
                const isStartOfPage = finalY === 0 || targetPage.boxes.length === 0;

                if (fits) {
                    placed = true;
                } else if (isStartOfPage) {
                    placed = true;
                } else {
                    const nextPageIndex = targetPage.pageIndex + 1;
                    let nextPage = currentPages.find(p => p.pageIndex === nextPageIndex);
                    if (!nextPage) {
                        nextPage = { pageIndex: nextPageIndex, boxes: [], w: materialWidth, h: 0 };
                        currentPages.push(nextPage);
                    }
                    targetPage = nextPage;
                    finalY = 0; 
                    finalX = 0;
                    currentInsertionPoint = null;
                }
            }

            const newItem = {
                ...item,
                id: isPreset ? `preset-${Date.now()}-${Math.random()}` : `${item.id}-${Date.now()}-${Math.random()}`,
                x: finalX,
                y: finalY,
                rotated: rotated,
                rotationAngle: rotated ? 90 : 0,
                pageIndex: targetPage.pageIndex,
                isSupplement: isPreset ? true : item.isSupplement
            };

            const newBoxes = [...targetPage.boxes, newItem];
            const newPageH = Math.max(targetPage.h, finalY + itemH);
            
            const pageIdx = currentPages.findIndex(p => p.pageIndex === targetPage.pageIndex);
            if (pageIdx !== -1) {
                currentPages[pageIdx] = { ...targetPage, boxes: newBoxes, h: newPageH };
            }

            // Update insertion point for next item (continue to the right)
            currentInsertionPoint = {
                pageIndex: targetPage.pageIndex,
                x: finalX + itemW + spacing,
                y: finalY
            };
        });

        const newTotalH = currentPages.reduce((acc, p) => acc + p.h, 0);
        setLayout({ ...layout, pages: currentPages, totalH: newTotalH });
        
        if (!isPreset) {
            const insertedIds = new Set(items.map(i => i.id));
            setMaterialBoxItems(prev => prev.filter(i => !insertedIds.has(i.id)));
            
            if (selectedBoxGroupKey) {
                 const remainingGroupItems = materialBoxItems.filter(i => !insertedIds.has(i.id) && `${i.internalOrderNumber}|${i.w}|${i.h}` === selectedBoxGroupKey);
                 if (remainingGroupItems.length === 0) {
                     setSelectedBoxGroupKey(null);
                     setSelectedBoxItems([]);
                 } else {
                     setSelectedBoxItems(remainingGroupItems);
                 }
            }
        }
    };

    const handleUndo = () => {
        if (layoutHistory.length === 0) return;
        const lastState = layoutHistory[layoutHistory.length - 1];
        setLayout(lastState.layout);
        setMaterialBoxItems(lastState.boxItems);
        setLayoutHistory(prev => prev.slice(0, -1));
    };

    const handleDrop = (e: React.DragEvent) => {
        e.preventDefault();
        const dataStr = e.dataTransfer.getData("application/json");
        if (!dataStr || !layout || !canvasRef.current || !canvasContainerRef.current) return;

        const { key, isPreset } = JSON.parse(dataStr);
        let itemsToInsert: AppItem[] = [];

        if (isPreset) {
            const presetIdx = parseInt(key.split('-')[1]);
            itemsToInsert = [JSON.parse(JSON.stringify(PRESET_SUPPLEMENTS[presetIdx]))];
        } else {
            // Find items in current materialBoxItems
            const groupKey = key;
            itemsToInsert = materialBoxItems.filter(i => `${i.internalOrderNumber}|${i.w}|${i.h}` === groupKey);
        }

        if (itemsToInsert.length === 0) return;

        const container = canvasContainerRef.current;
        const rect = canvasRef.current.getBoundingClientRect();
        const mouseX = e.clientX - rect.left;
        const mouseY = e.clientY - rect.top;

        const transformedX = (mouseX - canvasTransform.x) / canvasTransform.scale;
        const transformedY = (mouseY - canvasTransform.y) / canvasTransform.scale;

        const drawScale = container.clientWidth / layout.totalW;
        const lx = transformedX / drawScale;
        const ly = transformedY / drawScale;

        // Find which page was dropped on
        let cumulativeY = 50; // Canvas starting offset
        let targetPage = layout.pages[0];
        let relativeY = ly - 50;

        for (let i = 0; i < layout.pages.length; i++) {
            const page = layout.pages[i];
            const pageBottom = cumulativeY + page.h;
            // Detect if drop is within this page's band
            if (ly >= cumulativeY && ly <= pageBottom + (i < layout.pages.length - 1 ? SECTION_GAP : 1000)) {
                targetPage = page;
                relativeY = ly - cumulativeY;
                break;
            }
            cumulativeY += page.h + SECTION_GAP;
        }

        pushHistory(layout, materialBoxItems);

        const newPages = [...layout.pages];
        const pageIdx = newPages.findIndex(p => p.pageIndex === targetPage.pageIndex);
        
        let currentDropY = relativeY;
        const processedItems = itemsToInsert.map(item => {
            const rotated = (item.w > materialWidth && item.h <= materialWidth);
            const newItem = {
                ...item,
                id: isPreset ? `preset-${Date.now()}-${Math.random()}` : item.id,
                x: lx,
                y: currentDropY,
                rotated,
                rotationAngle: rotated ? 90 : 0,
                pageIndex: targetPage.pageIndex,
                isSupplement: isPreset ? true : item.isSupplement
            };
            const h = rotated ? item.w : item.h;
            currentDropY += 5; // Slight offset if multiple items dropped at once
            return newItem;
        });

        const updatedBoxes = [...newPages[pageIdx].boxes, ...processedItems];
        let newMaxH = 0;
        updatedBoxes.forEach(b => {
            const bh = b.rotated ? b.w : b.h;
            newMaxH = Math.max(newMaxH, b.y + bh);
        });

        newPages[pageIdx] = { ...newPages[pageIdx], boxes: updatedBoxes, h: newMaxH };
        const newTotalH = newPages.reduce((sum, p) => sum + p.h, 0);
        setLayout({ ...layout, pages: newPages, totalH: newTotalH });

        if (!isPreset) {
            const insertedIds = new Set(itemsToInsert.map(i => i.id));
            setMaterialBoxItems(prev => prev.filter(i => !insertedIds.has(i.id)));
        }
    };

    // Keyboard Handling
    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            if (['INPUT', 'TEXTAREA'].includes((e.target as HTMLElement).tagName)) return;

            // Z: Undo
            if (e.key.toLowerCase() === 'z' && !e.ctrlKey && !e.metaKey && (isMaterialBoxVisible || isPresetBoxVisible)) {
                handleUndo();
                return; 
            }

            // Box Item Operations
            if (isMaterialBoxVisible && selectedBoxItems.length > 0) {
                if (e.key.toLowerCase() === 'a') {
                    handleBoxItemInsert(selectedBoxItems, false);
                    return;
                }
                if (e.key.toLowerCase() === 's') {
                    handleBoxItemInsert(selectedBoxItems, true);
                    return;
                }
            }
            
            if (isPresetBoxVisible && selectedPresetItems.length > 0) {
                if (e.key.toLowerCase() === 'a') {
                    handleBoxItemInsert(selectedPresetItems, false, true);
                    return;
                }
                if (e.key.toLowerCase() === 's') {
                    handleBoxItemInsert(selectedPresetItems, true, true);
                    return;
                }
            }

            if (!layout) return;

            // Delete / Backspace (Canvas Items)
            if (e.key === 'Delete' || e.key === 'Backspace') {
                e.preventDefault();
                if (selectedItemIds.size === 0) return;

                const newPages = layout.pages.map(page => {
                    const newBoxes = page.boxes.filter(b => !selectedItemIds.has(b.id) || !b.isSupplement);
                    let maxY = 0;
                    newBoxes.forEach(p => {
                        const boxH = p.rotated ? p.w : p.h;
                        if (p.y + boxH > maxY) maxY = p.y + boxH;
                    });
                    return { ...page, boxes: newBoxes, h: maxY };
                });
                const newTotalH = newPages.reduce((sum, p) => sum + p.h, 0);
                setLayout({ ...layout, pages: newPages, totalH: newTotalH });
                setSelectedItem(null);
                setHighlightedItemId(null);
                setSelectedItemIds(new Set());
            }

            // Copy (Ctrl+C)
            if ((e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C')) {
                e.preventDefault();
                if (selectedItemIds.size === 0) return;

                const newPages = [...layout.pages];
                const newIds = new Set<string>();
                
                selectedItemIds.forEach(id => {
                    let pageIdx = -1;
                    let item: AppItem | undefined;
                    newPages.forEach((p, idx) => {
                        const found = p.boxes.find(b => b.id === id);
                        if (found) { pageIdx = idx; item = found; }
                    });

                    if (item && pageIdx !== -1) {
                        const page = { ...newPages[pageIdx] };
                        const currentHeight = item.rotated ? item.w : item.h;
                        
                        const newBox = {
                            ...item,
                            id: `${item.internalOrderNumber}-copy-${Date.now()}-${Math.random()}`,
                            originalId: item.originalId || item.id,
                            isSupplement: true, 
                            x: item.x,
                            y: item.y + currentHeight + spacing,
                            rotationAngle: item.rotationAngle || 0,
                            pageIndex: page.pageIndex
                        };
                        
                        const boxH = newBox.rotated ? newBox.w : newBox.h;
                        if (newBox.y + boxH > page.h) {
                             page.h = newBox.y + boxH;
                        }
                        
                        page.boxes = [...page.boxes, newBox];
                        newPages[pageIdx] = page;
                        newIds.add(newBox.id);
                    }
                });
                
                const newTotalH = newPages.reduce((sum, p) => sum + p.h, 0);
                setLayout({ ...layout, pages: newPages, totalH: newTotalH });
                setSelectedItemIds(newIds);
                if (newIds.size > 0) {
                    const firstId = newIds.values().next().value;
                    const firstBox = newPages.flatMap(p => p.boxes).find(b => b.id === firstId);
                    setSelectedItem(firstBox || null);
                }
            }
            
            // Rotate (r or R) (Canvas Items)
            if (e.key === 'r' || e.key === 'R') {
                e.preventDefault();
                if (selectedItemIds.size === 0) return;

                const pageGroups = new Map<number, Set<string>>();
                layout.pages.forEach((p, idx) => {
                    p.boxes.forEach(b => {
                        if (selectedItemIds.has(b.id)) {
                            if (!pageGroups.has(idx)) pageGroups.set(idx, new Set());
                            pageGroups.get(idx).add(b.id);
                        }
                    });
                });

                const newPages = layout.pages.map((page, pIdx) => {
                    if (!pageGroups.has(pIdx)) return page;
                    const idsOnPage = pageGroups.get(pIdx);
                    let pMinX = MAX_INT, pMinY = MAX_INT, pMaxX = -MAX_INT, pMaxY = -MAX_INT;
                    let count = 0;
                    page.boxes.forEach(b => {
                        if (idsOnPage.has(b.id)) {
                            count++;
                            const bW = b.rotated ? b.h : b.w;
                            const bH = b.rotated ? b.w : b.h;
                            if (b.x < pMinX) pMinX = b.x;
                            if (b.y < pMinY) pMinY = b.y;
                            if (b.x + bW > pMaxX) pMaxX = b.x + bW;
                            if (b.y + bH > pMaxY) pMaxY = b.y + bH;
                        }
                    });

                    if (count === 0) return page;
                    const cx = (pMinX + pMaxX) / 2;
                    const cy = (pMinY + pMaxY) / 2;
                    let pageModified = false;
                    const newBoxes = page.boxes.map(b => {
                        if (idsOnPage.has(b.id)) {
                            pageModified = true;
                            const curW = b.rotated ? b.h : b.w;
                            const curH = b.rotated ? b.w : b.h;
                            const ix = b.x + curW / 2;
                            const iy = b.y + curH / 2;
                            let newX, newY;
                            if (count > 1) {
                                const nix = cx - (iy - cy);
                                const niy = cy + (ix - cx);
                                const newW = curH; const newH = curW;
                                newX = nix - newW / 2; newY = niy - newH / 2;
                            } else {
                                const newW = curH; const newH = curW;
                                newX = ix - newW / 2; newY = iy - newH / 2;
                            }
                            const currentAngle = b.rotationAngle || (b.rotated ? 90 : 0);
                            const newAngle = (currentAngle + 90) % 360;
                            const newRotated = newAngle % 180 !== 0;
                            return { ...b, rotationAngle: newAngle, rotated: newRotated, x: newX, y: newY };
                        }
                        return b;
                    });
                    if (pageModified) {
                        let maxY = 0;
                        newBoxes.forEach(p => {
                             const boxH = p.rotated ? p.w : p.h;
                             if (p.y + boxH > maxY) maxY = p.y + boxH;
                        });
                        return { ...page, boxes: newBoxes, h: maxY };
                    }
                    return page;
                });
                const newTotalH = newPages.reduce((sum, p) => sum + p.h, 0);
                setLayout({ ...layout, pages: newPages, totalH: newTotalH });
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [layout, selectedItem, selectedItemIds, spacing, isMaterialBoxVisible, isPresetBoxVisible, selectedBoxItems, selectedPresetItems, layoutHistory]);

    const processFile = async () => {
        if (!file) { setError('请先上传订单文件 (Please upload an order file first).'); return; }
        setIsLoading(true); setError(''); setLayout(null); setGroupedOrders({}); setStagedItems([]); setUnprocessedItems([]); setCurrentStep(1);

        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                const headers = json[0] as string[];
                const orderNumberIndex = headers.indexOf('内部订单号');
                const productCodeIndex = headers.indexOf('商品编码');
                const notesIndex = headers.indexOf('订单备注');
                const quantityIndex = headers.indexOf('数量');
                
                if (orderNumberIndex === -1 || productCodeIndex === -1 || notesIndex === -1 || quantityIndex === -1) {
                  throw new Error('文件缺少必要的列：内部订单号, 商品编码, 订单备注, 数量');
                }

                const ordersToProcess = json.slice(1).map(row => ({
                    internalOrderNumber: row[orderNumberIndex]?.toString() || '',
                    productCode: row[productCodeIndex]?.toString() || '',
                    notes: row[notesIndex]?.toString() || '',
                    quantity: row[quantityIndex] ? parseInt(row[quantityIndex].toString(), 10) : 1
                })).filter(row => row.internalOrderNumber || row.productCode || row.notes);

                const orderDetailsMap = new Map();
                ordersToProcess.forEach(order => { if(order.internalOrderNumber) orderDetailsMap.set(order.internalOrderNumber, { productCode: order.productCode, notes: order.notes }); });
                setOrderDetails(orderDetailsMap);
                
                const { stagedItems: parsedStaged, unprocessedItems: initialUnprocessed } = parseOrderDataLocal(ordersToProcess, settings);
                const finalStaged = []; const itemsForReview = [];
                for (const item of parsedStaged) {
                    if (item.material === '未知材质' || item.color === '未知颜色') itemsForReview.push(item); else finalStaged.push(item);
                }
                
                const finalUnprocessed = [
                    ...initialUnprocessed.map(item => ({ internalOrderNumber: item.internalOrderNumber, details: item.details, })),
                    ...itemsForReview.map(item => ({ internalOrderNumber: item.internalOrderNumber, details: { productCode: item.productCode, notes: item.notes }, w_manual: item.w, h_manual: item.h, cornerRadius_manual: item.cornerRadius.tl, material_manual: item.material, color_manual: item.color, }))
                ];

                setStagedItems(finalStaged); setUnprocessedItems(finalUnprocessed);
                if (finalStaged.length > 0 || finalUnprocessed.length > 0) setCurrentStep(2); else setError("文件中未找到可处理的订单数据。 (No processable order data found in the file.)");

            } catch (err) { setError(`处理文件失败 (Failed to process file): ${err.message}`); } finally { setIsLoading(false); }
        };
        reader.readAsArrayBuffer(file);
    };
    
    const handleStagedItemChange = useCallback((index, field, value, cornerKey = null) => {
        setStagedItems(prevStagedItems => {
            const newStagedItems = [...prevStagedItems]; let item = { ...newStagedItems[index] };
            if (field === 'cornerRadius') {
                const numValue = parseInt(value, 10) || 0;
                if (cornerKey) item.cornerRadius = { ...item.cornerRadius, [cornerKey]: numValue }; else item.cornerRadius = { tl: numValue, tr: numValue, br: numValue, bl: numValue };
            } else if (['w', 'h', 'qty'].includes(field)) item[field] = parseInt(value, 10) || 0; else item[field] = value;
            if (field !== 'classification') item.classification = determineClassification(item, settings);
            newStagedItems[index] = item;
            if (editingReviewItem && editingReviewItem.index === index) setEditingReviewItem({ item: newStagedItems[index], index });
            return newStagedItems;
        });
    }, [editingReviewItem]);

    const handleDeleteStagedItem = useCallback((idToDelete) => {
        setStagedItems(prevStaged => {
            const itemToMove = prevStaged.find(item => item.id === idToDelete);
            if (itemToMove) {
                const unprocessedVersion = {
                    internalOrderNumber: itemToMove.internalOrderNumber, details: { productCode: itemToMove.productCode, notes: itemToMove.notes },
                    material_manual: itemToMove.material, color_manual: itemToMove.color, w_manual: itemToMove.w, h_manual: itemToMove.h, cornerRadius_manual: itemToMove.cornerRadius.tl,
                };
                setUnprocessedItems(prevUnprocessed => [unprocessedVersion, ...prevUnprocessed]);
            }
            return prevStaged.filter(item => item.id !== idToDelete);
        });
    }, []);

    const handleAddStagedItem = () => {
        const newItem: AppItem = { id: `manual-${Date.now()}`, internalOrderNumber: 'MANUAL', productCode: 'MANUAL', notes: 'MANual ENTRY', material: '默认材质', color: '默认颜色', w: 100, h: 100, qty: 1, cornerRadius: { tl: 0, tr: 0, br: 0, bl: 0 }, isCustom: true, classification: '定制', };
        setStagedItems(prev => [newItem, ...prev]);
    };

    const [tabSort, setTabSort] = useState<'alpha' | 'qty_desc' | 'qty_asc'>('alpha');
    const [selectedMaterialFilter, setSelectedMaterialFilter] = useState<string>('全部');

    const filteredGroupKeys = useMemo(() => {
        if (!groupedOrders) return [];
        return Object.keys(groupedOrders)
            .filter(key => {
                if (key === '全部混排') return false;
                if (selectedMaterialFilter === '全部') return true;
                return key.startsWith(selectedMaterialFilter + ' /');
            })
            .sort((a, b) => {
                if (tabSort === 'qty_desc') return groupedOrders[b].length - groupedOrders[a].length;
                if (tabSort === 'qty_asc') return groupedOrders[a].length - groupedOrders[b].length;
                return a.localeCompare(b, 'zh-CN');
            });
    }, [groupedOrders, selectedMaterialFilter, tabSort]);

    useEffect(() => {
        if (currentStep === 3 && filteredGroupKeys.length > 0) {
            if (!selectedGroupKey || (selectedGroupKey !== '全部混排' && !filteredGroupKeys.includes(selectedGroupKey))) {
                handleGroupSelection(filteredGroupKeys[0]);
            }
        }
    }, [filteredGroupKeys, currentStep]);

    const handleConfirmAndGroup = () => {
        const newGroups = {}; 
        const allItemsGroup = [];

        stagedItems.forEach(order => {
            if (order.material && order.color && order.w > 0 && order.h > 0) {
                const mixedKey = `${order.material} / ${order.color}`;
                
                if (!newGroups[mixedKey]) newGroups[mixedKey] = [];

                const quantity = parseInt(String(order.qty), 10) || 1;
                for (let i = 0; i < quantity; i++) {
                    const baseId = `${order.id}-${i}`;
                    const itemMixedGroup = { ...order, id: baseId, originalId: order.id, };
                    newGroups[mixedKey].push(itemMixedGroup);

                    const itemAllMixed = { ...order, id: `${baseId}-mix`, originalId: order.id, };
                    allItemsGroup.push(itemAllMixed);
                }
            }
        });

        if (allItemsGroup.length > 0) newGroups['全部混排'] = allItemsGroup;
        
        setGroupedOrders(newGroups);
        const firstGroupKey = Object.keys(newGroups).sort().find(k => k !== '全部混排') || '全部混排';
        setSelectedGroupKey(firstGroupKey);
        setSelectedMaterialFilter('全部');
        setLayout(null);
        setCurrentStep(3);
    };

    const stagedUniqueMaterials = useMemo(() => {
        const counts: Record<string, number> = {};
        stagedItems.forEach(item => {
            if (item.material) {
                const qty = parseInt(String(item.qty), 10) || 1;
                counts[item.material] = (counts[item.material] || 0) + qty;
            }
        });
        return Object.keys(counts).sort().map(name => ({ name, count: counts[name] }));
    }, [stagedItems]);

    const handleSyncRules = (itemIndex: number) => {
        const item = unprocessedItems[itemIndex];
        const material = item.material_manual;
        const color = item.color_manual;
        const notes = item.details?.notes || '';
        const productCode = item.details?.productCode || '';
        const fullText = `${productCode} ${notes}`;
        
        let updatedCount = 0;
        const newSettings = { ...settings };
        
        // 1. 同步材质 (排除掉通用的未知词)
        if (material && !['未知材质', '', 'N/A'].includes(material) && !newSettings.materialKeywords.includes(material)) {
            newSettings.materialKeywords = [...newSettings.materialKeywords, material];
            updatedCount++;
        }
        
        // 2. 同步颜色
        if (color && !['未知颜色', '', 'N/A'].includes(color) && !newSettings.colorKeywords.includes(color)) {
            newSettings.colorKeywords = [...newSettings.colorKeywords, color];
            updatedCount++;
        }

        // 3. 自动同步分类规律
        const commonClasses = ['异形', '定制', '中间椭圆', '椭圆', '圆形', '圆角'];
        for (const cls of commonClasses) {
            if (fullText.includes(cls)) {
                const exists = newSettings.classificationRules.some(r => r.keyword === cls);
                if (!exists) {
                    newSettings.classificationRules = [...newSettings.classificationRules, { keyword: cls, result: cls }];
                    updatedCount++;
                }
            }
        }
        
        if (updatedCount > 0) {
            setSettings(newSettings);
            localStorage.setItem('nesting_app_settings', JSON.stringify(newSettings));
            alert(`成功同步 ${updatedCount} 条新规则！下次导入将自动识别这些信息。`);
        } else {
            alert("未发现需要同步的新规则（已存在的规律无需重复同步）。");
        }
    };

    const handleAutoParseUnprocessed = (index: number, forcedSettings?: AppSettings) => {
        const item = unprocessedItems[index];
        if (!item) return;
        const notes = item.details?.notes || '';
        const productCode = item.details?.productCode || '';
        const currentSettings = forcedSettings || settings;
        
        const mockOrder = { internalOrderNumber: item.internalOrderNumber, productCode, notes, quantity: 1 };
        const result = parseOrderDataLocal([mockOrder], currentSettings);
        
        // 我们需要更多的智能：如果 productCode 或 notes 包含特定关键词，我们也补充给手动字段
        let foundMaterial = item.material_manual;
        let foundColor = item.color_manual;
        
        if (!foundMaterial || foundMaterial === '未知材质') {
            for (const m of currentSettings.materialKeywords) { if (`${productCode} ${notes}`.includes(m)) { foundMaterial = m; break; } }
        }
        if (!foundColor || foundColor === '未知颜色') {
            for (const c of currentSettings.colorKeywords) { if (`${productCode} ${notes}`.includes(c)) { foundColor = c; break; } }
        }

        const pStaged = result.stagedItems[0];
        const pUnprocessed = result.unprocessedItems[0];
        
        const parsed = pStaged ? {
            w: pStaged.w,
            h: pStaged.h,
            material: pStaged.material && pStaged.material !== '未知材质' ? pStaged.material : foundMaterial,
            color: pStaged.color && pStaged.color !== '未知颜色' ? pStaged.color : foundColor,
            cornerRadius: pStaged.cornerRadius
        } : (pUnprocessed ? {
            w: 0, h: 0,
            material: pUnprocessed.material_manual && pUnprocessed.material_manual !== '未知材质' ? pUnprocessed.material_manual : foundMaterial,
            color: pUnprocessed.color_manual && pUnprocessed.color_manual !== '未知颜色' ? pUnprocessed.color_manual : foundColor,
            cornerRadius: { tl: 0 }
        } : null);

        if (parsed) {
            const newUnprocessed = [...unprocessedItems];
            const old = newUnprocessed[index];
            
            newUnprocessed[index] = {
                ...old,
                w_manual: parsed.w > 0 ? parsed.w : old.w_manual,
                h_manual: parsed.h > 0 ? parsed.h : old.h_manual,
                material_manual: parsed.material || old.material_manual,
                color_manual: parsed.color || old.color_manual,
                cornerRadius_manual: (parsed.cornerRadius && parsed.cornerRadius.tl !== undefined && parsed.cornerRadius.tl > 0) ? parsed.cornerRadius.tl : old.cornerRadius_manual
            };
            
            setUnprocessedItems(newUnprocessed);
            const hasDataChange = JSON.stringify(newUnprocessed[index]) !== JSON.stringify(old);
            if (!hasDataChange) {
                alert("未发现新的信息。请尝试手动输入或检查设置中的尺寸连接符。");
            }
        } else {
            alert("解析失败，请检查设置中的尺寸连接符是否包含订单中使用的符号。");
        }
    };

    const handleUnprocessedItemChange = (index, field, value) => { const newUnprocessed = [...unprocessedItems]; newUnprocessed[index] = { ...newUnprocessed[index], [field]: value }; setUnprocessedItems(newUnprocessed); };

    const handleReProcessItem = (itemIndex) => {
        const itemToProcess = unprocessedItems[itemIndex];
        const width = parseFloat(itemToProcess.w_manual || 0); const height = parseFloat(itemToProcess.h_manual || 0);
        const radius = parseFloat(itemToProcess.cornerRadius_manual || 0);
        const material = itemToProcess.material_manual || ''; const color = itemToProcess.color_manual || '';
        if (width <= 0 || height <= 0) { alert("请输入有效的宽度和高度 (Please enter valid width and height)."); return; }
        if (!material || material === '未知材质' || !color || color === '未知颜色') { alert("请输入有效的物料和颜色 (Please enter valid material and color)."); return; }
        const newPieceData = { w: width, h: height, qty: 1, material, color, productCode: itemToProcess.details?.productCode || 'N/A', notes: itemToProcess.details?.notes || 'N/A', cornerRadius: { tl: radius, tr: radius, br: radius, bl: radius }, id: `${itemToProcess.internalOrderNumber}-manual-${Date.now()}`, internalOrderNumber: itemToProcess.internalOrderNumber, isCustom: (itemToProcess.details?.notes || '').includes('定制'), };
        const newPiece: AppItem = { ...newPieceData, classification: determineClassification(newPieceData, settings), }
        setStagedItems(prev => [...prev, newPiece]);
        const newUnprocessedItems = unprocessedItems.filter((_, index) => index !== itemIndex);
        setUnprocessedItems(newUnprocessedItems);
    };
    
    const handleGroupSelection = (key) => { setSelectedGroupKey(key); setLayout(null); setCanvasTransform({ scale: 1, x: 0, y: 0 }); setMaterialBoxItems([]); };
    
    const stopDeepSearch = () => { stopDeepSearchRef.current = true; };

    const separateItemsForBox = (items: AppItem[]) => {
        if (!selectedGroupKey || selectedGroupKey === '全部混排') {
            return { main: items, box: [] };
        }
        const counts = new Map<string, number>();
        items.forEach(it => counts.set(it.internalOrderNumber, (counts.get(it.internalOrderNumber) || 0) + 1));
        const mainItems: AppItem[] = [];
        const boxItems: AppItem[] = [];
        items.forEach(it => {
            if ((counts.get(it.internalOrderNumber) || 0) > 1 || it.classification === '异形') {
                boxItems.push(it);
            } else {
                mainItems.push(it);
            }
        });
        boxItems.sort((a, b) => a.internalOrderNumber.localeCompare(b.internalOrderNumber, 'zh-CN', { numeric: true }));
        return { main: mainItems, box: boxItems };
    };

    const runNesting = useCallback(async (customStrategy?: SortStrategy) => {
        if (!selectedGroupKey || !groupedOrders || !materialWidth || isNesting) return;
        stopDeepSearchRef.current = false; setIsNesting(true); setNestingProgress(0); setCanvasTransform({ scale: 1, x: 0, y: 0 });
        await new Promise(resolve => setTimeout(resolve, 10)); 
        try {
            const allItems = [...groupedOrders[selectedGroupKey]];
            if (!allItems || allItems.length === 0) { setLayout(null); setIsNesting(false); return; };
            const { main, box } = separateItemsForBox(allItems);
            setMaterialBoxItems(box);
            if (box.length > 0) setIsMaterialBoxVisible(true);
            const strategy = customStrategy || 'AREA_DESC';
            const currentResult = packLayout(main, materialWidth, materialLength, spacing, allowRotation, useRandom, strategy, 'BSSF', groupByOrder, nestingAlgorithm);
            learnFromLayout(currentResult, nestingAlgorithm, strategy, 'BSSF', allowRotation);
            setLayout(currentResult);
            setIsNesting(false);
        } catch (e) { setError(`排版算法出错 (Algorithm failed): ${e.message}`); setIsNesting(false); }
    }, [selectedGroupKey, groupedOrders, materialWidth, materialLength, spacing, allowRotation, isNesting, useRandom, groupByOrder]);

    const handleGeneticOptimization = useCallback(async () => {
        if (!selectedGroupKey || !groupedOrders || !materialWidth || isNesting || isOptimizing) return;
        setIsOptimizing(true); 
        setOptimizationProgress(0); 
        setStatusText('准备开始...');
        setBestStats({ count: 0, height: 0 }); 
        setLayout(null);
        stopDeepSearchRef.current = false;
        await new Promise(r => setTimeout(r, 50));
        try {
            const allItems = [...groupedOrders[selectedGroupKey]];
            const { main, box } = separateItemsForBox(allItems);
            setMaterialBoxItems(box);
            if (box.length > 0) setIsMaterialBoxVisible(true);
            let globalBestResult = null;
            const TOTAL_RUNS = 10;
            for (let run = 0; run < TOTAL_RUNS; run++) {
                if (stopDeepSearchRef.current) break;
                await runGA(
                    main, 
                    materialWidth, 
                    materialLength, 
                    spacing, 
                    allowRotation, 
                    groupByOrder,
                    (genProgress, currentRunBestLayout, currentRunBestCount) => {
                        const totalProgress = (run * 10) + (genProgress / 10);
                        setOptimizationProgress(Math.round(totalProgress));
                        const currentGen = Math.round((genProgress / 100) * 40); 
                        setStatusText(`第 ${run + 1}/${TOTAL_RUNS} 轮 - 进化代数 ${currentGen}/40`);
                        let isGlobalBest = false;
                        if (!globalBestResult) {
                            isGlobalBest = true;
                        } else {
                            if (currentRunBestCount > globalBestResult.count) {
                                isGlobalBest = true;
                            } else if (currentRunBestCount === globalBestResult.count) {
                                if (currentRunBestLayout.totalH < globalBestResult.height) {
                                    isGlobalBest = true;
                                }
                            }
                        }
                        if (isGlobalBest) {
                            globalBestResult = {
                                layout: currentRunBestLayout,
                                count: currentRunBestCount,
                                height: currentRunBestLayout.totalH
                            };
                            setBestStats({ count: globalBestResult.count, height: globalBestResult.height });
                            setLayout(globalBestResult.layout);
                        }
                    },
                    () => stopDeepSearchRef.current,
                    nestingAlgorithm
                );
            }
        } catch (e) { setError(`优化过程出错: ${e.message}`); } finally { setIsOptimizing(false); setStatusText(''); }
    }, [selectedGroupKey, groupedOrders, materialWidth, materialLength, spacing, allowRotation, isNesting, isOptimizing, groupByOrder]);

    const handleAutoOptimize = async () => {
        if (!selectedGroupKey || !groupedOrders || !materialWidth || isNesting || isOptimizing) return;
        setIsOptimizing(true); setOptimizationProgress(0); setBestStats({ count: 0, height: 0 }); setCanvasTransform({ scale: 1, x: 0, y: 0 }); setLayout(null);
        await new Promise(resolve => setTimeout(resolve, 50));
        try {
            const allItems = [...groupedOrders[selectedGroupKey]];
            const { main, box } = separateItemsForBox(allItems);
            setMaterialBoxItems(box);
            if (box.length > 0) setIsMaterialBoxVisible(true);
            let bestLayout = null;
            let bestStrategyParams = null;
            let minTotalHeight = Number.MAX_SAFE_INTEGER;
            let maxItemsPacked = 0;
            const strategies = [];
            
            // Add library strategies (Top 5)
            strategyLibrary.slice(0, 5).forEach(s => {
                strategies.push({ 
                    sort: s.sortStrategy, 
                    heur: s.heuristic, 
                    random: false, 
                    algorithm: s.algorithm, 
                    allowRotation: s.allowRotation,
                    isFromLib: true 
                });
            });

            if (groupByOrder) {
                strategies.push({ sort: 'ORDER_PRIORITY_OPTIMIZED', heur: 'BSSF', random: false });
                strategies.push({ sort: 'ORDER_GROUP', heur: 'BSSF', random: false });
                strategies.push({ sort: 'ORDER_GROUP', heur: 'BLSF', random: false });
                strategies.push({ sort: 'ORDER_GROUP', heur: 'BAF', random: false });
                for(let k = 0; k < 5; k++) strategies.push({ sort: 'ORDER_GROUP', heur: 'BSSF', random: true });
            } else {
                strategies.push({ sort: 'SMART_WIDTH_MATCH', heur: 'BSSF', random: false, forceNoRotation: true });
                strategies.push({ sort: 'SMART_WIDTH_MATCH', heur: 'BLSF', random: false, forceNoRotation: true });
                (['BSSF', 'BLSF', 'BAF']).forEach(heur => {
                    strategies.push({ sort: 'AREA_DESC', heur, random: false });
                    strategies.push({ sort: 'LONGSIDE_DESC', heur, random: false });
                });
                strategies.push({ sort: 'WIDTH_DESC', heur: 'BSSF', random: false });
                strategies.push({ sort: 'COMPLEMENTARY', heur: 'BSSF', random: false });
                strategies.push({ sort: 'PERIMETER_DESC', heur: 'BSSF', random: false });
                strategies.push({ sort: 'SHORTSIDE_DESC', heur: 'BSSF', random: false });
                for(let k = 0; k < 5; k++) strategies.push({ sort: 'RANDOM', heur: 'BSSF', random: true });
            }
            const totalIterations = strategies.length;
            for (let i = 0; i < totalIterations; i++) {
                if (stopDeepSearchRef.current) { stopDeepSearchRef.current = false; break; }
                const strat = strategies[i];
                const effectiveRotation = strat.isFromLib ? strat.allowRotation : (strat.forceNoRotation ? false : allowRotation);
                const effectiveAlg = strat.algorithm || nestingAlgorithm;
                const result = packLayout(main, materialWidth, materialLength, spacing, effectiveRotation, strat.random, strat.sort, strat.heur, groupByOrder, effectiveAlg);
                const totalPacked = result.pages.reduce((sum, p) => sum + p.boxes.length, 0);
                const currentTotalH = result.totalH;
                let isBetter = false;
                if (totalPacked > maxItemsPacked) {
                    isBetter = true;
                } else if (totalPacked === maxItemsPacked) {
                    if (currentTotalH < minTotalHeight) {
                        isBetter = true;
                    }
                }
                if (isBetter) {
                    maxItemsPacked = totalPacked;
                    minTotalHeight = currentTotalH;
                    bestLayout = result;
                    bestStrategyParams = {
                        alg: effectiveAlg,
                        sort: strat.sort,
                        heur: strat.heur,
                        rot: effectiveRotation
                    };
                    setBestStats({ count: maxItemsPacked, height: minTotalHeight });
                }
                setOptimizationProgress(Math.round(((i + 1) / totalIterations) * 100));
                if (i % 2 === 0) await new Promise(resolve => setTimeout(resolve, 0));
            }
            if (bestLayout) {
                setLayout(bestLayout);
                if (bestStrategyParams) {
                    learnFromLayout(bestLayout, bestStrategyParams.alg, bestStrategyParams.sort, bestStrategyParams.heur, bestStrategyParams.rot);
                }
            } else {
                runNesting();
            }
        } catch (e) { setError(`优化过程出错: ${e.message}`); } finally { setIsOptimizing(false); }
   };

    const getClickedItem = (layoutX, layoutY) => {
        if (!layout) return null;
        let cumulativeY = 0;
        for(let i=0; i < layout.pages.length; i++) {
            const page = layout.pages[i];
            const pageHeight = page.h;
            if (layoutY >= cumulativeY && layoutY <= cumulativeY + pageHeight) {
                const relativeY = layoutY - cumulativeY;
                for (let box of page.boxes) {
                    const boxW = box.rotated ? box.h : box.w;
                    const boxH = box.rotated ? box.w : box.h;
                    if (layoutX >= box.x && layoutX <= box.x + boxW && relativeY >= box.y && relativeY <= box.y + boxH) {
                        return box;
                    }
                }
            }
            cumulativeY += pageHeight + SECTION_GAP;
        }
        return null;
    };

    const handleCanvasDoubleClick = (event) => {
        if (!layout || !canvasRef.current || !canvasContainerRef.current || isPanning) return;
        const canvas = canvasRef.current; const container = canvasContainerRef.current;
        const rect = canvas.getBoundingClientRect();
        const clientX = (event.clientX - rect.left); const clientY = (event.clientY - rect.top);
        const transformedX = (clientX - canvasTransform.x) / canvasTransform.scale;
        const transformedY = (clientY - canvasTransform.y) / canvasTransform.scale;
        const drawScale = container.clientWidth / layout.totalW;
        const layoutX = transformedX / drawScale; const layoutY = transformedY / drawScale;
        const clickedBox = getClickedItem(layoutX, layoutY);
        if (clickedBox) { 
            setSelectedItem(clickedBox); 
            setEditedItem({ ...clickedBox, w: clickedBox.w, h: clickedBox.h }); 
            setIsEditModalOpen(true); 
            setHighlightedItemId(clickedBox.id); 
            setSelectedItemIds(new Set([clickedBox.id]));
        }
    };
    
    const handleWheel = useCallback((e: WheelEvent) => {
        e.preventDefault(); const canvas = canvasRef.current; if (!canvas) return;
        const rect = canvas.getBoundingClientRect();
        const mouseX = e.clientX - rect.left; const mouseY = e.clientY - rect.top;
        const zoom = 1 - e.deltaY * 0.001;
        setCanvasTransform(prevTransform => {
            const newScale = Math.max(0.1, Math.min(prevTransform.scale * zoom, 10));
            const newX = mouseX - (mouseX - prevTransform.x) * (newScale / prevTransform.scale);
            const newY = mouseY - (mouseY - prevTransform.y) * (newScale / prevTransform.scale);
            return { scale: newScale, x: newX, y: newY };
        });
    }, []);

    useEffect(() => { const container = canvasContainerRef.current; if (container && layout) { container.addEventListener('wheel', handleWheel, { passive: false }); return () => { if (container) container.removeEventListener('wheel', handleWheel); }; } }, [layout, handleWheel]);

    const handleMouseDown = (e) => {
        if (!layout || !canvasRef.current || !canvasContainerRef.current) return;
        const container = canvasContainerRef.current; const rect = canvasRef.current.getBoundingClientRect();
        const clientX = e.clientX - rect.left; const clientY = e.clientY - rect.top;
        const transformedX = (clientX - canvasTransform.x) / canvasTransform.scale; const transformedY = (clientY - canvasTransform.y) / canvasTransform.scale;
        const drawScale = container.clientWidth / layout.totalW;
        const layoutX = transformedX / drawScale; const layoutY = transformedY / drawScale;
        if (e.altKey && e.button === 0) {
            setIsBoxSelecting(true);
            setSelectionBox({ x: layoutX, y: layoutY, w: 0, h: 0 });
            setSelectedItemIds(new Set());
            setSelectedItem(null);
            return;
        }
        const clickedBox = getClickedItem(layoutX, layoutY);
        if (clickedBox && e.button === 0 && !e.shiftKey) { 
            let newSelection = new Set(selectedItemIds);
            if (!newSelection.has(clickedBox.id)) {
                newSelection = new Set([clickedBox.id]);
                setSelectedItemIds(newSelection);
                setSelectedItem(clickedBox);
                setHighlightedItemId(clickedBox.id);
            } else {
                setSelectedItem(clickedBox); 
            }
            setDraggingItemId(clickedBox.id);
             let visualBoxY = clickedBox.y;
             let cumulativeY = 0;
             for(let i=0; i < layout.pages.length; i++) {
                 const page = layout.pages[i];
                 if (page.pageIndex === clickedBox.pageIndex) {
                     visualBoxY += cumulativeY;
                     break;
                 }
                 cumulativeY += page.h + SECTION_GAP;
             }
             setDragOffset({ x: layoutX - clickedBox.x, y: layoutY - visualBoxY });
        } else { setIsPanning(true); setPanStart({ x: e.clientX, y: e.clientY }); }
    };

    const rebalancePages = (currentLayout: FullLayout, lengthLimit: number) => {
        if (lengthLimit <= 0) return currentLayout;
        const newPages: PageLayout[] = [];
        let pendingItems: AppItem[] = [];
        const processItemsForPage = (items: AppItem[], pageIndex: number, isNew: boolean): { page: PageLayout, overflow: AppItem[] } => {
            const keepItems: AppItem[] = [];
            const overflow: AppItem[] = [];
            items.sort((a, b) => a.y - b.y);
            for (const item of items) {
                if (item.y >= lengthLimit) {
                    overflow.push({ ...item, y: item.y - lengthLimit });
                } else {
                    keepItems.push(item);
                }
            }
            let maxH = 0;
            keepItems.forEach(i => {
                const ih = i.rotated ? i.w : i.h;
                maxH = Math.max(maxH, i.y + ih);
            });
            // Update pageIndex for items that are kept on this page
            const updatedKeepItems = keepItems.map(item => ({...item, pageIndex}));
            return {
                page: { pageIndex: pageIndex, boxes: updatedKeepItems, w: currentLayout.totalW, h: maxH },
                overflow: overflow
            };
        };
        let currentPageIndex = 1;
        for (const page of currentLayout.pages) {
            const itemsToProcess = [...page.boxes, ...pendingItems];
            pendingItems = []; 
            const result = processItemsForPage(itemsToProcess, currentPageIndex, false);
            newPages.push(result.page);
            pendingItems = result.overflow;
            currentPageIndex++;
        }
        let safetyCount = 0;
        while (pendingItems.length > 0 && safetyCount < 50) {
            const result = processItemsForPage(pendingItems, currentPageIndex, true);
            newPages.push(result.page);
            pendingItems = result.overflow;
            currentPageIndex++;
            safetyCount++;
        }
        const totalH = newPages.reduce((sum, p) => sum + p.h, 0);
        return { ...currentLayout, pages: newPages, totalH };
    };

    const handleMouseUp = () => { 
        setIsPanning(false); 
        setDraggingItemId(null); 
        setGuideLines([]); 
        setIsBoxSelecting(false);
        setSelectionBox(null);
        if (layout && materialLength > 0) {
            const newLayout = rebalancePages(layout, materialLength);
            setLayout(newLayout);
            learnFromLayout(newLayout, 'MAXRECTS', 'MANUAL', 'BSSF', allowRotation);
        }
    };
    
    const handleMouseMove = (e) => {
        if (isPanning) {
            const dx = e.clientX - panStart.x;
            const dy = e.clientY - panStart.y;
            setCanvasTransform(prev => ({ ...prev, x: prev.x + dx, y: prev.y + dy }));
            setPanStart({ x: e.clientX, y: e.clientY });
        } else if (isBoxSelecting && selectionBox && layout) {
            const container = canvasContainerRef.current;
            const drawScale = container.clientWidth / layout.totalW;
            const rect = canvasRef.current.getBoundingClientRect();
            const clientX = e.clientX - rect.left;
            const clientY = e.clientY - rect.top;
            const transformedX = (clientX - canvasTransform.x) / canvasTransform.scale;
            const transformedY = (clientY - canvasTransform.y) / canvasTransform.scale;
            const layoutX = transformedX / drawScale;
            const layoutY = transformedY / drawScale;
            const newW = layoutX - selectionBox.x;
            const newH = layoutY - selectionBox.y;
            const currentBox = {
                x: newW > 0 ? selectionBox.x : selectionBox.x + newW,
                y: newH > 0 ? selectionBox.y : selectionBox.y + newH,
                w: Math.abs(newW),
                h: Math.abs(newH)
            };
            setSelectionBox({ ...selectionBox, w: newW, h: newH });
            const newSelectedIds = new Set<string>();
            let currentYOffset = 0;
            layout.pages.forEach(page => {
                page.boxes.forEach(box => {
                    const boxW = box.rotated ? box.h : box.w;
                    const boxH = box.rotated ? box.w : box.h;
                    const boxAbsX = box.x;
                    const boxAbsY = box.y + currentYOffset;
                    if (currentBox.x < boxAbsX + boxW && currentBox.x + currentBox.w > boxAbsX && currentBox.y < boxAbsY + boxH && currentBox.y + currentBox.h > boxAbsY) {
                        newSelectedIds.add(box.id);
                    }
                });
                currentYOffset += page.h + SECTION_GAP;
            });
            setSelectedItemIds(newSelectedIds);
        } else if (draggingItemId && layout && canvasContainerRef.current) {
            const container = canvasContainerRef.current;
            const drawScale = container.clientWidth / layout.totalW;
            const rect = canvasRef.current.getBoundingClientRect();
            const clientX = e.clientX - rect.left;
            const clientY = e.clientY - rect.top;
            const transformedX = (clientX - canvasTransform.x) / canvasTransform.scale;
            const transformedY = (clientY - canvasTransform.y) / canvasTransform.scale;
            const layoutX = transformedX / drawScale;
            const layoutY = transformedY / drawScale;
            let activePage = null;
            let leaderBox = null;
            layout.pages.forEach(p => {
                const found = p.boxes.find(b => b.id === draggingItemId);
                if (found) { activePage = p; leaderBox = found; }
            });
            if (activePage && leaderBox) {
                let pageVisualTop = 0;
                for(let i=0; i < layout.pages.length; i++) {
                    const page = layout.pages[i];
                    if (page.pageIndex === activePage.pageIndex) break;
                    pageVisualTop += page.h + SECTION_GAP;
                }
                let newX = layoutX - dragOffset.x;
                let newY = layoutY - dragOffset.y - pageVisualTop; 
                const SNAP_THRESHOLD = 5; 
                let snappedX = newX;
                let snappedY = newY;
                const myW = leaderBox.rotated ? leaderBox.h : leaderBox.w;
                const myH = leaderBox.rotated ? leaderBox.w : leaderBox.h;
                const guides = [];
                const xCandidates = [0, materialWidth - myW];
                const yCandidates = [0];
                
                activePage.boxes.forEach(other => {
                    if (selectedItemIds.has(other.id)) return; 
                    const otherW = other.rotated ? other.h : other.w;
                    const otherH = other.rotated ? other.w : other.h;
                    
                    const ox1 = other.x;
                    const ox2 = other.x + otherW;
                    const oy1 = other.y;
                    const oy2 = other.y + otherH;

                    // Add candidates only if they are potentially within snap range to save memory/loops
                    if (Math.abs(ox1 - newX) < SNAP_THRESHOLD || Math.abs(ox2 - myW - newX) < SNAP_THRESHOLD || 
                        Math.abs(ox2 - newX) < SNAP_THRESHOLD || Math.abs(ox1 - myW - newX) < SNAP_THRESHOLD ||
                        Math.abs(ox2 + spacing - newX) < SNAP_THRESHOLD || Math.abs(ox1 - myW - spacing - newX) < SNAP_THRESHOLD) {
                        xCandidates.push(ox1, ox2 - myW, ox2, ox1 - myW, ox2 + spacing, ox1 - myW - spacing);
                    }

                    if (Math.abs(oy1 - newY) < SNAP_THRESHOLD || Math.abs(oy2 - myH - newY) < SNAP_THRESHOLD ||
                        Math.abs(oy2 - newY) < SNAP_THRESHOLD || Math.abs(oy1 - myH - newY) < SNAP_THRESHOLD ||
                        Math.abs(oy2 + spacing - newY) < SNAP_THRESHOLD || Math.abs(oy1 - myH - spacing - newY) < SNAP_THRESHOLD) {
                        yCandidates.push(oy1, oy2 - myH, oy2, oy1 - myH, oy2 + spacing, oy1 - myH - spacing);
                    }
                });
                let minDiffX = MAX_INT;
                xCandidates.forEach(val => { const diff = Math.abs(val - newX); if (diff < SNAP_THRESHOLD && diff < minDiffX) { minDiffX = diff; snappedX = val; } });
                let minDiffY = MAX_INT;
                yCandidates.forEach(val => { const diff = Math.abs(val - newY); if (diff < SNAP_THRESHOLD && diff < minDiffY) { minDiffY = diff; snappedY = val; } });
                if (minDiffX < MAX_INT) guides.push({x1: snappedX + (layoutX - newX), y1: 0, x2: snappedX + (layoutX - newX), y2: layout.totalH}); 
                if (minDiffY < MAX_INT) guides.push({x1: 0, y1: snappedY + pageVisualTop, x2: layout.totalW, y2: snappedY + pageVisualTop});
                setGuideLines(guides);
                const dx = snappedX - leaderBox.x;
                const dy = snappedY - leaderBox.y;
                const newPages = layout.pages.map(p => {
                    let pageModified = false;
                    const updatedBoxes = p.boxes.map(b => {
                        if (selectedItemIds.has(b.id)) {
                            pageModified = true;
                            return { ...b, x: b.x + dx, y: b.y + dy };
                        }
                        return b;
                    });
                    if (pageModified) {
                        let maxY = 0;
                        updatedBoxes.forEach(box => {
                             const boxH = box.rotated ? box.w : box.h;
                             if (box.y + boxH > maxY) maxY = box.y + boxH;
                        });
                        return { ...p, boxes: updatedBoxes, h: maxY };
                    }
                    return p;
                });
                setLayout({ ...layout, pages: newPages, totalH: newPages.reduce((sum,p)=>sum+p.h,0) });
            }
        }
    };

    const duplicateMap = useMemo(() => {
        const dMap = new Map();
        allBoxes.forEach(b => {
            const num = b.internalOrderNumber;
            dMap.set(num, (dMap.get(num) || 0) + 1);
        });
        return dMap;
    }, [allBoxes]);

    const handleItemEdit = useCallback((field, value) => { if (!editedItem) return; setEditedItem(prev => ({ ...prev, [field]: parseFloat(value) || 0 })); }, [editedItem]);
    const handleCornerEdit = useCallback((corner, value) => { if (!editedItem) return; setEditedItem(prev => ({ ...prev, cornerRadius: { ...prev.cornerRadius, [corner]: parseFloat(value) || 0 } })); }, [editedItem]);
    const handleRowClick = useCallback((box) => { 
        setSelectedItem(box); 
        setHighlightedItemId(box.id); 
        setSelectedItemIds(new Set([box.id]));
    }, []);
    const resetCanvasView = useCallback(() => { setCanvasTransform({ scale: 1, x: 0, y: 0 }); }, []);
    const saveItemChanges = useCallback(() => {
        if (!editedItem || !layout) return;
        const newPages = layout.pages.map(page => ({
            ...page,
            boxes: page.boxes.map(box => box.id === editedItem.id ? { ...box, w: editedItem.w, h: editedItem.h, cornerRadius: editedItem.cornerRadius } : box)
        }));
        setLayout(prev => prev ? ({ ...prev, pages: newPages }) : null); setSelectedItem(editedItem); setIsEditModalOpen(false);
    }, [editedItem, layout]);

    const handleListDoubleClick = (box) => { setSelectedItem(box); setEditedItem({ ...box, w: box.w, h: box.h }); setIsEditModalOpen(true); setHighlightedItemId(box.id); setSelectedItemIds(new Set([box.id])); };
    useEffect(() => {
        if (!layout || !canvasRef.current || !canvasContainerRef.current) return;
        let animationFrameId;
        
        const render = () => {
            const canvas = canvasRef.current; 
            const container = canvasContainerRef.current; 
            if (!canvas || !container) return;
            const ctx = canvas.getContext('2d');
            const parentWidth = container.clientWidth;
            const scale = parentWidth / layout.totalW;
            let totalCanvasHeightLayoutUnits = 0;
            layout.pages.forEach((p, i) => {
                 totalCanvasHeightLayoutUnits += p.h;
                 if (i < layout.pages.length - 1) totalCanvasHeightLayoutUnits += SECTION_GAP;
            });
            totalCanvasHeightLayoutUnits += 100;
            
            // Only resize if needed
            const newW = layout.totalW * scale;
            const newH = totalCanvasHeightLayoutUnits * scale;
            if (canvas.width !== newW) canvas.width = newW;
            if (canvas.height !== newH) canvas.height = newH;

            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.save();
            ctx.translate(canvasTransform.x, canvasTransform.y);
            ctx.scale(canvasTransform.scale, canvasTransform.scale);
            ctx.save();
            ctx.scale(scale, scale);
            const effectiveScale = scale * canvasTransform.scale;
            
            // Calculate viewport in layout units
            const viewportX = -canvasTransform.x / effectiveScale;
            const viewportY = -canvasTransform.y / effectiveScale;
            const viewportW = canvas.width / effectiveScale;
            const viewportH = canvas.height / effectiveScale;

            let currentYOffset = 50; 

            // Cache Path2D objects on items if missing
            allBoxes.forEach(box => {
                if (box.pathData && !box._path2d) {
                    try {
                        (box as any)._path2d = new Path2D(box.pathData);
                    } catch (e) {
                        console.error("Invalid path data", e);
                    }
                }
            });

            layout.pages.forEach((page, pIdx) => {
                const pageHeight = page.h;
                const boundaryH = materialLength > 0 ? materialLength : pageHeight;
                
                // Draw page boundary only if visible
                const pageTop = currentYOffset;
                const pageBottom = currentYOffset + Math.max(boundaryH, pageHeight);
                
                const isPageVisible = !(pageBottom < viewportY || pageTop > viewportY + viewportH);

                if (isPageVisible) {
                    ctx.strokeStyle = '#999';
                    ctx.lineWidth = 2 / effectiveScale;
                    ctx.setLineDash([10 / effectiveScale, 10 / effectiveScale]);
                    ctx.strokeRect(0, currentYOffset, layout.totalW, boundaryH);
                    ctx.setLineDash([]);
                    ctx.fillStyle = '#6c757d';
                    ctx.font = `bold ${24 / effectiveScale}px Arial`;
                    ctx.textAlign = 'left';
                    ctx.textBaseline = 'bottom';
                    ctx.fillText(`Section ${page.pageIndex} (Used: ${pageHeight}mm / Limit: ${boundaryH}mm)`, 0, currentYOffset - (5 / effectiveScale));

                    // Optimize overlap detection: if too many items, limit or skip
                    const overlaps = new Set();
                    if (page.boxes.length < 300) { // Reduced O(N^2) guard further
                        for (let i = 0; i < page.boxes.length; i++) {
                            const b1 = page.boxes[i];
                            const b1W = b1.rotated ? b1.h : b1.w;
                            const b1H = b1.rotated ? b1.w : b1.h;
                            if (currentYOffset + b1.y + b1H < viewportY || currentYOffset + b1.y > viewportY + viewportH) continue;

                            for (let j = i + 1; j < page.boxes.length; j++) {
                                const b2 = page.boxes[j];
                                const b2W = b2.rotated ? b2.h : b2.w;
                                const b2H = b2.rotated ? b2.w : b2.h;
                                if (!(b1.x + b1W <= b2.x || b1.x >= b2.x + b2W || b1.y + b1H <= b2.y || b1.y >= b2.y + b2H)) {
                                    overlaps.add(b1.id); overlaps.add(b2.id);
                                }
                            }
                        }
                    }

                    page.boxes.forEach(box => {
                        const drawW = box.rotated ? box.h : box.w;
                        const drawH = box.rotated ? box.w : box.h;
                        const absX = box.x;
                        const absY = box.y + currentYOffset;

                        // Viewport culling
                        if (absX + drawW < viewportX || absX > viewportX + viewportW || 
                            absY + drawH < viewportY || absY > viewportY + viewportH) {
                            return;
                        }

                        const isHighlighted = highlightedItemId === box.id || selectedItemIds.has(box.id);
                        const isHovered = hoveredItemId === box.id;
                        const isIrregular = box.classification === '异形';
                        const isOverlapping = overlaps.has(box.id);
                        const isOutOfBounds = (box.x < 0) || (box.x + drawW > layout.totalW); 
                        
                        let fillColor;
                        if (isOverlapping || isOutOfBounds) fillColor = 'rgba(220, 53, 69, 0.5)';
                        else if (isIrregular) fillColor = 'rgba(108, 117, 125, 0.2)';
                        else if (box.isCustom) fillColor = 'rgba(253, 126, 20, 0.2)';
                        else if ((duplicateMap.get(box.internalOrderNumber) || 0) > 1 || box.qty > 1) {
                            const baseColor = stringToColor(box.internalOrderNumber);
                            const r = parseInt(baseColor.slice(1, 3), 16);
                            const g = parseInt(baseColor.slice(3, 5), 16);
                            const b = parseInt(baseColor.slice(5, 7), 16);
                            fillColor = `rgba(${r}, ${g}, ${b}, 0.6)`;
                        } else fillColor = 'rgba(0, 123, 255, 0.1)';

                        const angle = box.rotationAngle !== undefined ? box.rotationAngle : (box.rotated ? 90 : 0);
                        ctx.save();
                        ctx.translate(absX + drawW / 2, absY + drawH / 2);
                        ctx.rotate(angle * Math.PI / 180);
                        ctx.translate(-box.w / 2, -box.h / 2);
                        ctx.fillStyle = fillColor;
                        
                        if (isHighlighted) { ctx.strokeStyle = '#28a745'; ctx.lineWidth = 4 / effectiveScale; }
                        else if (isHovered) { ctx.strokeStyle = '#ffc107'; ctx.lineWidth = 2 / effectiveScale; }
                        else { ctx.strokeStyle = isIrregular ? 'rgba(108, 117, 125, 0.8)' : (box.isCustom ? 'rgba(253, 126, 20, 0.8)' : 'rgba(0, 123, 255, 0.8)'); ctx.lineWidth = 1 / effectiveScale; }
                        
                        if (isOverlapping || isOutOfBounds) { ctx.strokeStyle = '#dc3545'; ctx.lineWidth = 3 / effectiveScale; }

                        if (box.pathData) { 
                            const path = (box as any)._path2d || new Path2D(box.pathData);
                            ctx.fill(path); 
                            ctx.stroke(path); 
                        } else { 
                            drawRoundedRect(ctx, 0, 0, box.w, box.h, box.cornerRadius); 
                            ctx.fill(); 
                            ctx.stroke(); 
                        }
                        
                        if (!isIrregular) {
                            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
                            const centerX = box.w / 2; const centerY = box.h / 2;
                            let text1 = box.isSupplement ? '补数' : box.internalOrderNumber;
                            let text2 = `${box.w}x${box.h}`;
                            const smallerDim = Math.min(box.w, box.h); const idealFontSize = (smallerDim * 0.3) / 2;
                            ctx.font = `bold ${idealFontSize}px Arial`;
                            const textWidth = ctx.measureText(text2).width; const maxTextWidth = box.w * 0.9;
                            let finalFontSize = idealFontSize; if (textWidth > maxTextWidth) finalFontSize = idealFontSize * (maxTextWidth / textWidth);
                            
                            if (finalFontSize * effectiveScale > 5) {
                                ctx.fillStyle = '#333'; const textLineHeight = finalFontSize * 1.2;
                                ctx.font = `bold ${finalFontSize}px Arial`;
                                ctx.fillText(text1 || '', centerX, centerY - textLineHeight / 2); 
                                ctx.fillText(text2, centerX, centerY + textLineHeight / 2);
                            }
                        }
                        ctx.restore();
                    });
                }

                currentYOffset += pageHeight;
                if (pIdx < layout.pages.length - 1) {
                    const isGapVisible = !(currentYOffset + SECTION_GAP < viewportY || currentYOffset > viewportY + viewportH);
                    if (isGapVisible) {
                        ctx.fillStyle = 'rgba(40, 167, 69, 0.15)'; 
                        ctx.fillRect(0, currentYOffset, layout.totalW, SECTION_GAP);
                        ctx.fillStyle = '#28a745';
                        ctx.font = `bold ${40 / effectiveScale}px Arial`;
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(`50cm Gap`, layout.totalW / 2, currentYOffset + SECTION_GAP / 2);
                    }
                    currentYOffset += SECTION_GAP;
                }
            });
            if (selectionBox) {
                const { x, y, w, h } = selectionBox;
                ctx.fillStyle = 'rgba(0, 123, 255, 0.3)';
                ctx.strokeStyle = '#007bff';
                ctx.lineWidth = 1 / effectiveScale;
                ctx.fillRect(x, y, w, h);
                ctx.strokeRect(x, y, w, h);
            }
            if (guideLines.length > 0) {
                ctx.strokeStyle = '#00ff00';
                ctx.lineWidth = 1 / effectiveScale;
                ctx.setLineDash([4 / effectiveScale, 2 / effectiveScale]);
                guideLines.forEach(line => { ctx.beginPath(); ctx.moveTo(line.x1, line.y1); ctx.lineTo(line.x2, line.y2); ctx.stroke(); });
                ctx.setLineDash([]);
            }
            ctx.restore(); ctx.restore();
        };

        animationFrameId = requestAnimationFrame(render);
        return () => cancelAnimationFrame(animationFrameId);
    }, [layout, selectedItem, highlightedItemId, hoveredItemId, canvasTransform, draggingItemId, materialWidth, materialLength, guideLines, materialBoxItems, selectionBox, selectedItemIds, duplicateMap]);

    const duplicateOrderNumbers = useMemo(() => { const counts = new Map(); stagedItems.forEach(item => counts.set(item.internalOrderNumber, (counts.get(item.internalOrderNumber) || 0) + 1)); const duplicates = new Set(); counts.forEach((count, key) => { if (count > 1) duplicates.add(key); }); return duplicates; }, [stagedItems]);
    
    const multiPieceItemIds = useMemo(() => {
        const ids = new Set<string>();
        stagedItems.forEach(item => {
            if ((item.qty || 1) > 1 || duplicateOrderNumbers.has(item.internalOrderNumber) || item.classification === '异形') {
                ids.add(item.id);
            }
        });
        return ids;
    }, [stagedItems, duplicateOrderNumbers]);

    const layoutDuplicateOrderNumbers = useMemo(() => { if (!layout) return new Set(); const counts = new Map(); allBoxes.forEach(box => { if (box.internalOrderNumber) counts.set(box.internalOrderNumber, (counts.get(box.internalOrderNumber) || 0) + 1); }); const duplicates = new Set(); counts.forEach((count, key) => { if (count > 1) duplicates.add(key); }); return duplicates; }, [layout, allBoxes]);
    const totalSheetCount = useMemo(() => stagedItems.reduce((acc, item) => acc + (item.qty || 0), 0), [stagedItems]);

    const layoutStats = useMemo(() => {
        if (!layout) return null;
        const totalArea = allBoxes.reduce((acc, box) => acc + (box.w * box.h), 0);
        let orderQty = 0; let orderArea = 0; let suppQty = 0; let suppArea = 0;
        allBoxes.forEach(box => {
            const area = box.w * box.h;
            if (box.isSupplement) { suppQty++; suppArea += area; } else { orderQty++; orderArea += area; }
        });
        const usedArea = layout.totalH * layout.totalW; 
        const wasteArea = usedArea - totalArea;
        const fill = usedArea > 0 ? totalArea / usedArea : 0;
        return { h: layout.totalH, totalArea, usedArea, wasteArea, fill, orderQty, orderArea, suppQty, suppArea, packedQty: allBoxes.length };
    }, [layout, allBoxes]);

    const exportToPlt = useCallback(() => {
        if (!layout) return;
        const now = new Date(); const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        let combinedPltContent = 'IN;VS30;SP1;\n';
        let currentYOffset = 0;
        let totalLayoutHeightWithGaps = 0;
        layout.pages.forEach((p, i) => { totalLayoutHeightWithGaps += p.h; if (i < layout.pages.length - 1) totalLayoutHeightWithGaps += SECTION_GAP; });
        layout.pages.forEach((page, index) => {
             page.boxes.forEach(box => {
                const adjustedY = box.y + currentYOffset;
                const angle = box.rotationAngle !== undefined ? box.rotationAngle : (box.rotated ? 90 : 0);
                if (box.pathData) combinedPltContent += pathDataToPlt(box.pathData, box.x, adjustedY, totalLayoutHeightWithGaps, angle, box.w, box.h);
                else combinedPltContent += generatePltPathForRect(box.w, box.h, box.cornerRadius, box.x, adjustedY, totalLayoutHeightWithGaps, angle);
             });
             currentYOffset += page.h; if (index < layout.pages.length - 1) currentYOffset += SECTION_GAP;
        });
        combinedPltContent += 'PU;SP0;IN;\n';
        const safeGroupKey = selectedGroupKey ? selectedGroupKey.replace(/[\\/:*?"<>|]/g, '_') : 'Layout';
        const filename = `${dateStr}-${safeGroupKey}-${layoutStats?.packedQty || 0}件.plt`;
        const blob = new Blob([combinedPltContent], { type: 'text/plain;charset=utf-8' });
        const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = filename; link.click(); URL.revokeObjectURL(link.href);
    }, [layout, selectedGroupKey, layoutStats]);

    const calculateItemCost = useCallback((w: number, h: number, material: string) => {
        const setting = unitCosts[material];
        if (!setting || setting.cost === 0) return 0;
        if (setting.unit === 'm') {
            const materialWidth = setting.width || 0;
            const maxDim = Math.max(w, h);
            const minDim = Math.min(w, h);
            if (materialWidth > 0) {
                if (maxDim > materialWidth) return (maxDim / 1000) * setting.cost;
                else return (minDim / 1000) * setting.cost;
            }
            return 0;
        } else {
            return (w * h / 1000000) * (typeof setting === 'number' ? setting : setting.cost);
        }
    }, [unitCosts]);

    const sortedDisplayItems = useMemo(() => {
        const sortedBoxes = [...allBoxes].sort((a,b) => (a.pageIndex || 0) - (b.pageIndex || 0) || a.y - b.y || a.x - b.x);
        if (mergeOrderNumbers) {
            const orderGroups = new Map<string, any>();
            const supplements: any[] = [];
            sortedBoxes.forEach(box => {
                if (box.isSupplement) {
                    supplements.push(box);
                    return;
                }
                if (!box.internalOrderNumber) {
                    supplements.push(box);
                    return;
                }
                const cost = calculateItemCost(box.w, box.h, box.material);
                if (!orderGroups.has(box.internalOrderNumber)) {
                    orderGroups.set(box.internalOrderNumber, { ...box, totalCost: cost });
                } else {
                    const existing = orderGroups.get(box.internalOrderNumber);
                    existing.totalCost += cost;
                }
            });
            return [...Array.from(orderGroups.values()), ...supplements].sort((a,b) => (a.pageIndex || 0) - (b.pageIndex || 0) || a.y - b.y || a.x - b.x);
        }
        return sortedBoxes;
    }, [allBoxes, mergeOrderNumbers, calculateItemCost]);

    const exportSortedListToExcel = useCallback(() => {
        if (!layout) return;
        const sortedBoxes = [...allBoxes].sort((a, b) => { 
            if (a.pageIndex !== b.pageIndex) return (a.pageIndex || 1) - (b.pageIndex || 1); 
            return (a.y - b.y) || (a.x - b.x); 
        });

        let finalDisplayItems: any[] = [];
        if (mergeOrderNumbers) {
            const orderGroups = new Map<string, any>();
            sortedBoxes.forEach(box => {
                if (box.isSupplement) return;
                if (!box.internalOrderNumber) return;
                const cost = calculateItemCost(box.w, box.h, box.material);
                if (!orderGroups.has(box.internalOrderNumber)) {
                    orderGroups.set(box.internalOrderNumber, { ...box, totalCost: cost });
                } else {
                    const existing = orderGroups.get(box.internalOrderNumber);
                    existing.totalCost += cost;
                }
            });
            finalDisplayItems = Array.from(orderGroups.values());
        } else {
            finalDisplayItems = sortedBoxes.filter(box => !box.isSupplement);
        }

        let regularCount = 0;
        const data = finalDisplayItems.map((item) => {
            const details = orderDetails.get(item.internalOrderNumber);
            let idxDisplay = '-'; let orderNumDisplay = '-'; let productCodeDisplay = '';
            const rawNotes = item.notes || (details ? details.notes : 'N/A'); 
            regularCount++; 
            idxDisplay = regularCount.toString().padStart(3, '0'); 
            orderNumDisplay = item.internalOrderNumber; 
            productCodeDisplay = details ? details.productCode : 'N/A'; 
            let notesDisplay = `${idxDisplay}.${rawNotes}`;
            
            let costVal = 0;
            if (mergeOrderNumbers) {
                costVal = item.totalCost;
            } else {
                costVal = calculateItemCost(item.w, item.h, item.material);
            }
            const cost = costVal.toFixed(2);

            const row = {};
            if (selectedExportColumns.includes('index')) row['序号'] = idxDisplay;
            if (selectedExportColumns.includes('internalOrderNumber')) row['内部订单号'] = orderNumDisplay;
            if (selectedExportColumns.includes('productCode')) row['商品编码'] = productCodeDisplay;
            if (selectedExportColumns.includes('notes')) row['卖家备注'] = notesDisplay;
            if (selectedExportColumns.includes('material')) row['材质'] = item.material;
            if (selectedExportColumns.includes('color')) row['颜色'] = item.color;
            if (selectedExportColumns.includes('w')) row['宽(mm)'] = item.w;
            if (selectedExportColumns.includes('h')) row['高(mm)'] = item.h;
            if (selectedExportColumns.includes('rotated')) row['旋转'] = item.rotated ? '是' : '否';
            if (selectedExportColumns.includes('cost')) row['成本价'] = cost;
            if (selectedExportColumns.includes('page')) row['分段'] = item.pageIndex;
            return row;
        });
        const now = new Date(); const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        const safeGroupKey = selectedGroupKey ? selectedGroupKey.replace(/[\\/:*?"<>|]/g, '_') : 'Layout';
        const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "排版清单");
        XLSX.writeFile(wb, `${dateStr}-${safeGroupKey}-${finalDisplayItems.length}件.xlsx`);
    }, [allBoxes, layout, selectedGroupKey, orderDetails, selectedExportColumns, unitCosts, mergeOrderNumbers]);

    const exportStatsToExcel = useCallback(() => {
        if (!layoutStats || !selectedGroupKey) return;
        const now = new Date(); const dateStr = now.toLocaleDateString() + ' ' + now.toLocaleTimeString();
        const data = [{
            '日期时间': dateStr, '材质/颜色': selectedGroupKey, '总页数': layout.pages.length,
            '已排版数量': layoutStats.packedQty, '订单数量': layoutStats.orderQty, '订单面积(m²)': (layoutStats.orderArea / 1_000_000).toFixed(3),
            '补数数量': layoutStats.suppQty, '补数面积(m²)': (layoutStats.suppArea / 1_000_000).toFixed(3), '利用率': (layoutStats.fill * 100).toFixed(2) + '%',
            '总长度(mm)': layoutStats.h.toFixed(0), '总面积(m²)': (layoutStats.totalArea / 1_000_000).toFixed(3),
            '物料面积(m²)': (layoutStats.usedArea / 1_000_000).toFixed(3), '浪费面积(m²)': (layoutStats.wasteArea / 1_000_000).toFixed(3)
        }];
        const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "排版统计");
        XLSX.writeFile(wb, `排版统计-${selectedGroupKey}-${Date.now()}.xlsx`);
    }, [layoutStats, selectedGroupKey, layout]);

    const classificationTypes = useMemo(() => {
        const base = ['全部', '多件分类'];
        const fromSettings = Array.from(new Set(settings.classificationRules.map(r => r.result).filter(Boolean)));
        const standard = ['常规', '圆角', '圆形', '椭圆', '中间椭圆', '异形', '定制'];
        return Array.from(new Set([...base, ...standard, ...fromSettings]));
    }, [settings.classificationRules]);

    const filteredStagedItems = useMemo(() => {
        if (activeFilter === '全部') return stagedItems;
        if (activeFilter === '多件分类') return stagedItems.filter(item => duplicateOrderNumbers.has(item.internalOrderNumber) || item.qty > 1);
        
        return stagedItems.filter(item => {
            const tags = (item.classification || '').split(',');
            return tags.includes(activeFilter);
        });
    }, [stagedItems, activeFilter, duplicateOrderNumbers]);

    return (
        <>
            <header>
                <h1>智能排版助手</h1>
                <p>上传您的订单表，自动提取信息、优化排版并导出切割文件。</p>
                <div className="header-nav">
                    <button className={`nav-btn ${!isSettingsOpen ? 'active' : ''}`} onClick={() => setIsSettingsOpen(false)}>主页 (Home)</button>
                    <button className={`nav-btn ${isSettingsOpen ? 'active' : ''}`} onClick={() => setIsSettingsOpen(true)}>设置 (Settings)</button>
                </div>
            </header>
            <main>
                {isSettingsOpen ? (
                    <SettingsPage 
                        localSettings={localSettings} 
                        setLocalSettings={setLocalSettings} 
                        onUpdate={setSettings} 
                        onBack={() => setIsSettingsOpen(false)} 
                    />
                ) : (
                    <div className="container">
                    <div className="card">
                        <h2>第一步：上传与配置</h2>
                        <div className="form-row">
                            <div className="form-group"><label htmlFor="file-upload">上传订单文件 (.xlsx, .csv)</label><input id="file-upload" type="file" accept=".xlsx, .csv" onChange={(e) => setFile(e.target.files[0])} /></div>
                            <div className="form-group"><label htmlFor="material-width">物料宽度 (mm)</label><input id="material-width" type="number" value={materialWidth} onChange={(e) => setMaterialWidth(parseInt(e.target.value, 10))} /></div>
                            <div className="form-group"><label htmlFor="material-length">物料长度 (mm, 0为无限)</label><input id="material-length" type="number" value={materialLength} onChange={(e) => setMaterialLength(parseInt(e.target.value, 10))} /></div>
                            <div className="form-group"><label htmlFor="spacing">排版间距 (mm)</label><input id="spacing" type="number" value={spacing} onChange={(e) => setSpacing(parseInt(e.target.value, 10))} /></div>
                        </div>
                        <div className="button-group"><button onClick={processFile} disabled={isLoading || !file} className="button button-primary">{isLoading ? `处理中...` : '开始处理'}</button><button onClick={resetState} className="button" style={{backgroundColor: '#6c757d', color: 'white'}}>重置</button></div>
                        {error && <p className="error-message">{error}</p>}
                    </div>
                    {currentStep >= 2 && unprocessedItems.length > 0 && (
                        <div className="unprocessed-items-container">
                             <div className="unprocessed-header"><h3>未处理订单 - 手动校正</h3></div>
                            <p>以下订单因信息无法识别或被删除而待处理。请在下方输入正确尺寸后，点击按钮将其添加到下方审核表格中。</p>
                            <table className="unprocessed-items-table">
                                <thead><tr><th>内部订单号</th><th>商品编码</th><th>订单备注</th><th>手动输入</th><th>操作</th></tr></thead>
                                <tbody>
                                    {unprocessedItems.map((item, index) => (
                                        <tr key={`${item.internalOrderNumber}-${index}`}>
                                            <td style={{wordBreak: 'break-all', whiteSpace: 'normal', minWidth: '80px'}}>{item.internalOrderNumber}</td>
                                            <td style={{wordBreak: 'break-all', whiteSpace: 'normal', minWidth: '120px'}}>{item.details?.productCode || 'N/A'}</td>
                                            <td style={{wordBreak: 'break-all', whiteSpace: 'normal', minWidth: '200px', maxWidth: '400px'}}>{item.details?.notes || 'N/A'}</td>
                                            <td className="manual-input-cell">
                                                <div style={{display: 'flex', gap: '8px', marginBottom: '8px', alignItems: 'center'}}>
                                                    <span style={{fontWeight: 'bold', color: '#856404'}}>手动输入</span>
                                                    <button className="button button-small" onClick={() => setMappingModal({ open: true, itemIndex: index })} style={{padding: '2px 8px', fontSize: '12px', background: '#e7f5ff', color: '#1971c2'}}>映射设置</button>
                                                </div>
                                                <div className="input-grid">
                                                    <input type="text" placeholder="材质" value={item.material_manual || ''} onChange={(e) => handleUnprocessedItemChange(index, 'material_manual', e.target.value)} />
                                                    <input type="text" placeholder="颜色" value={item.color_manual || ''} onChange={(e) => handleUnprocessedItemChange(index, 'color_manual', e.target.value)} />
                                                    <input type="number" placeholder="宽度(mm)" value={item.w_manual || ''} onChange={(e) => handleUnprocessedItemChange(index, 'w_manual', e.target.value)} />
                                                    <input type="number" placeholder="高度(mm)" value={item.h_manual || ''} onChange={(e) => handleUnprocessedItemChange(index, 'h_manual', e.target.value)} />
                                                    <input type="number" placeholder="圆角(mm)" value={item.cornerRadius_manual || ''} onChange={(e) => handleUnprocessedItemChange(index, 'cornerRadius_manual', e.target.value)} />
                                                </div>
                                            </td>
                                            <td>
                                                <div style={{display: 'flex', flexDirection: 'column', gap: '8px'}}>
                                                    <button className="button button-primary" onClick={() => handleReProcessItem(index)}>添加到审核列表</button>
                                                    <button className="button" style={{fontSize: '12px', padding: '4px', background: '#f8f9fa', border: '1px solid #ddd'}} onClick={() => handleSyncRules(index)}>同步规则 (基于此行补充规则)</button>
                                                </div>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                    {currentStep >= 2 && (stagedItems.length > 0 || unprocessedItems.length === 0) && (
                         <div className={`card ${isReviewFullScreen ? 'fullscreen' : ''}`}>
                            <div className="card-header">
                                <h2>第二步：审核与编辑数据</h2>
                                <div className="filter-buttons">
                                    {classificationTypes.map(type => (<button key={type} className={`button button-small ${activeFilter === type ? 'active' : ''}`} onClick={() => setActiveFilter(type)}>{type}</button>))}
                                </div>
                                <button onClick={() => setIsReviewFullScreen(!isReviewFullScreen)} className="button">{isReviewFullScreen ? '退出全屏' : '全屏'}</button>
                            </div>
                            <div className="card-description"><p>检查提取的数据。您可以在此表中手动修改、删除或添加新行。确认无误后，点击下方按钮进行分组排版。</p><p className="total-count"><strong>统计总数量:</strong> {totalSheetCount}</p></div>
                            <div className="review-table-container">
                                <table className="review-table">
                                    <thead><tr><th>序号</th><th>内部订单号</th><th>预览图</th><th>商品编码</th><th>订单备注</th><th>分类</th><th>数量</th><th>材质</th><th>颜色</th><th>宽度 (mm)</th><th>高度 (mm)</th><th>圆角 (mm)</th><th>操作</th></tr></thead>
                                    <tbody>
                                        {filteredStagedItems.map((item, index) => (
                                            <StagedItemRow key={item.id} item={item} rowIndex={index + 1} isDuplicate={duplicateOrderNumbers.has(item.internalOrderNumber)} index={stagedItems.findIndex(st => st.id === item.id)} onChange={handleStagedItemChange} onDelete={handleDeleteStagedItem} onPreviewClick={() => setEditingReviewItem({ item, index: stagedItems.findIndex(st => st.id === item.id) })} onMouseEnter={() => setHoveredOrderNumber(item.internalOrderNumber)} onMouseLeave={() => setHoveredOrderNumber(null)} isHighlighted={hoveredOrderNumber === item.internalOrderNumber} />
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                            <div className="button-group"><button onClick={handleAddStagedItem} className="button">添加新行</button><button onClick={handleConfirmAndGroup} className="button button-primary" disabled={stagedItems.length === 0}>确认并分组进行排版</button></div>
                        </div>
                    )}

                    {currentStep === 3 && groupedOrders && Object.keys(groupedOrders).length > 0 && (
                        <div className="card">
                             <h2>第三步：选择物料组进行排版</h2>
                             <div className="button-group" style={{marginBottom: '1rem'}}><button onClick={() => setCurrentStep(2)} className="button">返回编辑数据</button></div>
                             
                             <div className="tab-controls" style={{
                                 display: 'flex', 
                                 flexDirection: 'column',
                                 gap: '1rem', 
                                 marginBottom: '1rem', 
                                 padding: '1rem', 
                                 backgroundColor: '#f8f9fa', 
                                 borderRadius: '8px',
                                 border: '1px solid #e9ecef'
                             }}>
                                 <div style={{display: 'flex', gap: '0.5rem', alignItems: 'center', flexWrap: 'wrap'}}>
                                     <span style={{fontSize: '0.9rem', fontWeight: 'bold', color: '#495057'}}>材质筛选:</span>
                                     <button 
                                         className={`button button-small ${selectedMaterialFilter === '全部' ? 'active' : ''}`} 
                                         style={selectedMaterialFilter === '全部' ? { background: '#1971c2', color: '#fff' } : {}}
                                         onClick={() => setSelectedMaterialFilter('全部')}
                                     >
                                         全部预览 ({stagedItems.reduce((acc, curr) => acc + (parseInt(String(curr.qty), 10) || 1), 0)})
                                     </button>
                                     {stagedUniqueMaterials.map(mat => (
                                         <button 
                                             key={mat.name}
                                             className={`button button-small ${selectedMaterialFilter === mat.name ? 'active' : ''}`} 
                                             style={selectedMaterialFilter === mat.name ? { background: '#2f9e44', color: '#fff' } : {}}
                                             onClick={() => setSelectedMaterialFilter(mat.name)}
                                         >
                                             {mat.name} ({mat.count})
                                         </button>
                                     ))}
                                     <button 
                                         className={`button button-small ${selectedGroupKey === '全部混排' ? 'button-primary' : 'button-secondary'}`}
                                         style={{marginLeft: 'auto'}}
                                         onClick={() => handleGroupSelection('全部混排')}
                                     >
                                         ⚡ 全部混排 ({groupedOrders['全部混排']?.length || 0})
                                     </button>
                                 </div>
                                 <div style={{display: 'flex', gap: '0.5rem', alignItems: 'center'}}>
                                     <span style={{fontSize: '0.9rem', fontWeight: 'bold', color: '#495057'}}>组内排序:</span>
                                     <button className={`button button-small ${tabSort === 'alpha' ? 'active' : ''}`} onClick={() => setTabSort('alpha')}>拼音</button>
                                     <button className={`button button-small ${tabSort === 'qty_desc' ? 'active' : ''}`} onClick={() => setTabSort('qty_desc')}>数量 ↓</button>
                                     <button className={`button button-small ${tabSort === 'qty_asc' ? 'active' : ''}`} onClick={() => setTabSort('qty_asc')}>数量 ↑</button>
                                 </div>
                             </div>

                            <div className="group-tabs">
                                {filteredGroupKeys.map(key => {
                                    const groupItems = groupedOrders[key];
                                    const totalCount = groupItems.length;
                                    const multiPieceCount = groupItems.filter(item => multiPieceItemIds.has(item.originalId)).length;
                                    return (
                                        <button key={key} className={`button group-tab ${selectedGroupKey === key ? 'active' : ''}`} onClick={() => handleGroupSelection(key)}>
                                            {key} ({totalCount} 件)
                                            {multiPieceCount > 0 && (
                                                <span style={{ color: selectedGroupKey === key ? 'lightgreen' : 'var(--success-color)', marginLeft: '0.5rem', fontWeight: 'bold' }}>
                                                    - 多件{multiPieceCount}
                                                </span>
                                            )}
                                        </button>
                                    );
                                })}
                            </div>
                            {selectedGroupKey && (
                                <div className="layout-controls">
                                     {isOptimizing ? (
                                        <div className="optimization-status" style={{flex: 1, display: 'flex', flexDirection: 'column', gap: '0.5rem'}}>
                                            <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}><p style={{margin: 0, fontWeight: 'bold', color: '#ffc107'}}>{statusText || `AI 正在进化... ${optimizationProgress}%`}</p>{bestStats.count > 0 && (<span style={{fontSize: '0.9rem', color: '#28a745'}}>当前最佳: {bestStats.count} 件, 高度 {bestStats.height} mm</span>)}</div>
                                            <div style={{width: '100%', height: '10px', backgroundColor: '#e9ecef', borderRadius: '5px', overflow: 'hidden'}}><div style={{width: `${optimizationProgress}%`, height: '100%', backgroundColor: '#6f42c1', transition: 'width 0.1s linear'}}></div></div>
                                        </div>
                                     ) : (
                                        <>
                                            <label className="inline" style={{marginRight: '1rem', cursor: 'pointer', userSelect: 'none'}}><input type="checkbox" checked={useRandom} onChange={(e) => setUseRandom(e.target.checked)} /> 随机优化</label>
                                            <label className="inline" style={{marginRight: '1rem', cursor: 'pointer', userSelect: 'none'}}><input type="checkbox" checked={groupByOrder} onChange={(e) => setGroupByOrder(e.target.checked)} /> 相同订单相邻 (Order Grouping)</label>
                                            <select 
                                                id="nestingAlgorithm-action" 
                                                value={nestingAlgorithm} 
                                                onChange={(e) => setNestingAlgorithm(e.target.value as AlgorithmType)}
                                                style={{ marginRight: '1rem', padding: '6px 10px', borderRadius: '4px', border: '1px solid #ddd', backgroundColor: '#fff', fontSize: '0.9rem', fontWeight: 'bold', color: '#333', cursor: 'pointer' }}
                                            >
                                                <option value="MAXRECTS">核心: MaxRects (经典/最稳)</option>
                                                <option value="SKYLINE">核心: Skyline (更快/省料)</option>
                                                <option value="SHELF">核心: Shelf (极速/行列)</option>
                                            </select>
                                            <button onClick={() => setIsLibOpen(true)} className="button" style={{marginRight: '1rem', backgroundColor: '#4b5563', color: 'white'}}>📋 策略逻辑库 (Library)</button>
                                            <button onClick={() => runNesting('AREA_DESC')} className="button button-primary" style={{marginRight: '1rem'}}>开始排版 (Start Nesting)</button>
                                            <button onClick={handleGeneticOptimization} className="button button-secondary" style={{marginRight: '1rem', backgroundColor: '#6f42c1', borderColor: '#6f42c1'}}>🚀 AI 深度遗传算法排版 (GenAI Nesting)</button>
                                            <button onClick={handleAutoOptimize} className="button button-warning">💡 传统自动优化 (Auto-Find Best)</button>
                                        </>
                                     )}
                                </div>
                            )}
                            {layout && (
                                <>
                                    <div className={`results-grid ${isCanvasFullScreen ? 'fullscreen-canvas-grid' : ''}`}>
                                        <div 
                                            ref={canvasContainerRef} 
                                            className={`layout-container ${isCanvasFullScreen ? 'fullscreen' : ''}`} 
                                            onMouseDown={handleMouseDown} 
                                            onMouseMove={handleMouseMove} 
                                            onMouseUp={handleMouseUp} 
                                            onMouseLeave={handleMouseUp}
                                            onDragOver={(e) => e.preventDefault()}
                                            onDrop={handleDrop}
                                            onContextMenu={(e) => { e.preventDefault(); setSelectedItemIds(new Set()); setSelectedItem(null); setHighlightedItemId(null); }}
                                        >
                                            <canvas ref={canvasRef} id="nesting-canvas" onDoubleClick={handleCanvasDoubleClick}></canvas>
                                        </div>
                                        <div className="stats-panel">
                                            <div className="card">
                                                <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', borderBottom: '1px solid #eee', paddingBottom: '0.5rem'}}>
                                                    <h3 style={{margin: 0, border: 'none'}}>排版统计</h3>
                                                    <button onClick={() => setIsPresetBoxVisible(!isPresetBoxVisible)} className="button" style={{padding: '5px 10px', fontSize: '0.85rem', backgroundColor: '#6f42c1', color: 'white', border: 'none'}}>预设补数</button>
                                                </div>
                                                 <p style={{fontSize: '0.8rem', color: '#6c757d', marginBottom: '0.5rem'}}>Tips: 拖动时自动吸附, Ctrl+C 复制(补数), Backspace 删除(仅补数), R 旋转, Alt+拖动 框选, 右键 取消选择</p>
                                                 <button onClick={resetCanvasView} className="button" style={{width: '100%', marginBottom: '0.5rem'}}>重置视图</button>
                                                 <button onClick={() => setIsCanvasFullScreen(!isCanvasFullScreen)} className="button" style={{width: '100%', marginBottom: '1rem'}}>{isCanvasFullScreen ? '退出全屏' : '全屏查看'}</button>
                                                 {materialBoxItems.length > 0 && selectedGroupKey !== '全部混排' && (
                                                     <button onClick={() => setIsMaterialBoxVisible(!isMaterialBoxVisible)} className="button" style={{width: '100%', marginBottom: '1rem', border: '2px solid #007bff', color: '#007bff'}}>📦 多件物料框 ({materialBoxItems.length})</button>
                                                 )}
                                                {layoutStats && (
                                                <ul className="stats-list">
                                                    <li><span>已排版数量:</span> <strong>{layoutStats.packedQty}</strong></li>
                                                    <li><span>订单数量:</span> <strong>{layoutStats.orderQty}</strong></li>
                                                    <li><span>订单面积:</span> <strong>{(layoutStats.orderArea / 1_000_000).toFixed(3)} m²</strong></li>
                                                    <li><span>补数数量:</span> <strong>{layoutStats.suppQty}</strong></li>
                                                    <li><span>补数面积:</span> <strong>{(layoutStats.suppArea / 1_000_000).toFixed(3)} m²</strong></li>
                                                    <li><span>利用率:</span> <strong>{(layoutStats.fill * 100).toFixed(2)}%</strong></li>
                                                    <li><span>总长度:</span> <strong>{layoutStats.h.toFixed(0)} mm</strong></li>
                                                    <li><span>总面积:</span> <strong>{(layoutStats.totalArea / 1_000_000).toFixed(3)} m²</strong></li>
                                                    <li><span>物料面积:</span> <strong>{(layoutStats.usedArea / 1_000_000).toFixed(3)} m²</strong></li>
                                                    <li><span>浪费面积:</span> <strong>{(layoutStats.wasteArea / 1_000_000).toFixed(3)} m²</strong></li>
                                                </ul>
                                                )}
                                                <button onClick={exportStatsToExcel} className="button button-secondary" style={{width: '100%', marginTop: '1rem'}}>导出统计数据 (XLSX)</button>
                                            </div>
                                        </div>
                                    </div>
                                    <div className="form-group" style={{marginTop: '1.5rem'}}>
                                        <h2>第四步：导出结果</h2>
                                        <div className="cost-configuration" style={{marginBottom: '1rem', padding: '1rem', backgroundColor: '#fdfdfd', borderRadius: '8px', border: '1px solid #e9ecef'}}>
                                            <h4 style={{marginTop: 0, marginBottom: '0.8rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '1rem', color: '#495057'}}>
                                                材质成本单价设置 (自动同步至设置页)
                                                <span style={{fontSize: '0.75rem', fontWeight: 'normal', color: '#6c757d'}}>单位: 元/m² 或 元/米</span>
                                            </h4>
                                            <div style={{display: 'flex', flexDirection: 'column', gap: '0.8rem'}}>
                                                {uniqueMaterials.map(mat => {
                                                    const setting = unitCosts[mat] || { cost: 0, unit: 'sqm' };
                                                    return (
                                                        <div key={mat} style={{display: 'flex', gap: '10px', alignItems: 'center', borderBottom: '1px solid #f0f0f0', paddingBottom: '0.5rem', flexWrap: 'wrap'}}>
                                                            <label style={{fontSize: '0.85rem', color: '#495057', minWidth: '100px', fontWeight: 'bold'}}>{mat}</label>
                                                            <select 
                                                               value={setting.unit || 'sqm'} 
                                                               onChange={e => updateMaterialCostSetting(mat, 'unit', e.target.value)} 
                                                               style={{fontSize: '0.8rem', padding: '2px 4px', height: '28px', border: '1px solid #dee2e6', borderRadius: '4px'}}
                                                            >
                                                                <option value="sqm">按平方</option>
                                                                <option value="m">按米</option>
                                                            </select>
                                                            <div style={{display: 'flex', gap: '4px', alignItems: 'center'}}>
                                                                <span style={{fontSize: '0.75rem', color: '#888'}}>单价:</span>
                                                                <input 
                                                                   type="number" 
                                                                   value={setting.cost !== undefined ? setting.cost : ''} 
                                                                   onChange={e => handleUnitCostChange(mat, e.target.value)} 
                                                                   placeholder="0.00"
                                                                   step="0.1"
                                                                   style={{padding: '0.2rem 0.4rem', fontSize: '0.85rem', width: '80px', borderColor: '#e9ecef'}}
                                                               />
                                                            </div>
                                                            {setting.unit === 'm' && (
                                                                <div style={{display: 'flex', gap: '4px', alignItems: 'center'}}>
                                                                    <span style={{fontSize: '0.75rem', color: '#888'}}>宽幅:</span>
                                                                    <input 
                                                                       type="number" 
                                                                       value={setting.width || ''} 
                                                                       onChange={e => updateMaterialCostSetting(mat, 'width', parseFloat(e.target.value))} 
                                                                       placeholder="mm"
                                                                       style={{padding: '0.2rem 0.4rem', fontSize: '0.85rem', width: '80px', borderColor: '#e9ecef'}}
                                                                   />
                                                                </div>
                                                            )}
                                                        </div>
                                                    );
                                                })}
                                            </div>
                                        </div>
                                        <div style={{marginBottom: '1rem'}}>
                                            <label style={{marginBottom: '0.5rem', display: 'block'}}>选择导出字段:</label>
                                            <div style={{display: 'flex', gap: '1rem', flexWrap: 'wrap', alignItems: 'center'}}>
                                                {EXPORT_COLUMNS.map(col => (<label key={col.key} style={{display: 'flex', alignItems: 'center', gap: '0.3rem', cursor: 'pointer'}}><input type="checkbox" checked={selectedExportColumns.includes(col.key)} onChange={(e) => { if (e.target.checked) setSelectedExportColumns(prev => [...prev, col.key]); else setSelectedExportColumns(prev => prev.filter(k => k !== col.key)); }} />{col.label}</label>))}
                                                <label style={{display: 'flex', alignItems: 'center', gap: '0.3rem', cursor: 'pointer', marginLeft: '1rem', padding: '2px 8px', backgroundColor: '#fff4e6', borderRadius: '4px', border: '1px solid #ffd8a8', fontSize: '0.85rem', fontWeight: '500'}}>
                                                    <input type="checkbox" checked={mergeOrderNumbers} onChange={(e) => setMergeOrderNumbers(e.target.checked)} />
                                                    合并多件单号与成本
                                                </label>
                                            </div>
                                        </div>
                                        <div className="button-group"><button onClick={exportToPlt} className="button button-secondary">导出合并排版图 (PLT - Single File)</button><button onClick={exportSortedListToExcel} className="button button-secondary">导出排序清单 (XLSX)</button></div>
                                    </div>
                                    <h3>排版顺序清单</h3>
                                    <div style={{maxHeight: '400px', overflowY: 'auto'}}>
                                        <table className="sorted-orders-table">
                                            <thead><tr><th>分段</th><th>序号</th><th>内部订单号</th><th>商品编码</th><th>卖家备注</th><th>成本价</th></tr></thead>
                                            <tbody>
                                                {(() => {
                                                    let regularCount = 0;
                                                    return sortedDisplayItems.map((item) => {
                                                        const details = orderDetails.get(item.internalOrderNumber);
                                                        let idxDisplay = '-'; let orderNumDisplay = '-'; let productCodeDisplay = '';
                                                        const rawNotes = item.notes || (details ? details.notes : 'N/A'); 
                                                        if (item.isSupplement) { 
                                                            productCodeDisplay = item.internalOrderNumber || `补数-${item.w}x${item.h}`; 
                                                        } else { 
                                                            regularCount++; 
                                                            idxDisplay = regularCount.toString().padStart(3, '0'); 
                                                            orderNumDisplay = item.internalOrderNumber; 
                                                            productCodeDisplay = details ? details.productCode : 'N/A'; 
                                                        }
                                                        const isDuplicate = layoutDuplicateOrderNumbers.has(item.internalOrderNumber) && !item.isSupplement;
                                                        
                                                        let costDisplay = '0.00';
                                                        if (mergeOrderNumbers && !item.isSupplement && item.internalOrderNumber) {
                                                            costDisplay = (item.totalCost || 0).toFixed(2);
                                                        } else {
                                                            const unitCost = unitCosts[item.material] || 0; 
                                                            const areaM2 = (item.w * item.h) / 1000000; 
                                                            costDisplay = (areaM2 * (unitCost || 0)).toFixed(2);
                                                        }

                                                        return (
                                                            <tr key={item.id} data-item-id={item.id} onClick={() => handleRowClick(item)} onDoubleClick={() => handleListDoubleClick(item)} onMouseEnter={() => setHoveredItemId(item.id)} onMouseLeave={() => setHoveredItemId(null)} className={`${highlightedItemId === item.id ? 'selected' : ''} ${isDuplicate ? 'duplicate-row' : ''}`}>
                                                                <td>{item.pageIndex}</td><td>{idxDisplay}</td><td>{orderNumDisplay}</td><td>{productCodeDisplay}</td><td>{idxDisplay}.{rawNotes}</td><td>{costDisplay}</td>
                                                            </tr>
                                                        );
                                                    });
                                                })()}
                                            </tbody>
                                        </table>
                                    </div>
                                </>
                            )}
                        </div>
                    )}
                </div>
                )}
            </main>
            {isEditModalOpen && editedItem && (
                 <div className="modal-backdrop" onClick={() => setIsEditModalOpen(false)}>
                    <div className="modal-content" onClick={e => e.stopPropagation()}>
                        <div className="modal-header"><h4>编辑物件: {editedItem.isSupplement ? `补数-${editedItem.w}x${editedItem.h}` : editedItem.internalOrderNumber}</h4></div>
                        <div className="modal-body">
                            <div className="form-group"><label>宽度 (mm)</label><input type="number" value={editedItem.w} onChange={e => handleItemEdit('w', e.target.value)} /></div>
                            <div className="form-group"><label>高度 (mm)</label><input type="number" value={editedItem.h} onChange={e => handleItemEdit('h', e.target.value)} /></div>
                            <div className="form-group"><label>圆角半径 (mm)</label><div className="corner-inputs"><input type="number" title="Top-Left" placeholder="TL" value={editedItem.cornerRadius.tl} onChange={e => handleCornerEdit('tl', e.target.value)} /><input type="number" title="Top-Right" placeholder="TR" value={editedItem.cornerRadius.tr} onChange={e => handleCornerEdit('tr', e.target.value)} /><input type="number" title="Bottom-Left" placeholder="BL" value={editedItem.cornerRadius.bl} onChange={e => handleCornerEdit('bl', e.target.value)} /><input type="number" title="Bottom-Right" placeholder="BR" value={editedItem.cornerRadius.br} onChange={e => handleCornerEdit('br', e.target.value)} /></div></div>
                        </div>
                        <div className="modal-footer"><button onClick={() => setIsEditModalOpen(false)} className="button" style={{backgroundColor: '#6c757d', color: 'white'}}>取消</button><button onClick={saveItemChanges} className="button button-primary">保存并重排</button></div>
                    </div>
                 </div>
            )}
            {editingReviewItem && (<ReviewEditModal item={editingReviewItem.item} index={editingReviewItem.index} onChange={handleStagedItemChange} onClose={() => setEditingReviewItem(null)} />)}
            {multiInsertConfig.open && (
                <MultiInsertModal 
                    items={multiInsertConfig.items} 
                    onConfirm={(count, orientation) => {
                        handleBoxItemInsert(multiInsertConfig.items, null, false, count, orientation);
                        setMultiInsertConfig({ open: false, items: [] });
                    }}
                    onClose={() => setMultiInsertConfig({ open: false, items: [] })}
                />
            )}
            {isMaterialBoxVisible && materialBoxItems.length > 0 && selectedGroupKey !== '全部混排' && (
                <MaterialBox 
                    title="多件物料框"
                    items={materialBoxItems} 
                    onSelect={(key, items) => { setSelectedBoxGroupKey(key); setSelectedBoxItems(items); }}
                    selectedGroupKey={selectedBoxGroupKey}
                    onDoubleClick={(items) => handleBoxItemInsert(items)}
                    onClose={() => setIsMaterialBoxVisible(false)}
                    setMultiInsertConfig={setMultiInsertConfig}
                />
            )}
            {isPresetBoxVisible && (
                <MaterialBox
                    title="预设补数"
                    items={PRESET_SUPPLEMENTS as AppItem[]}
                    isPreset={true}
                    onSelect={(key, items) => { setSelectedPresetGroupKey(key); setSelectedPresetItems(items); }}
                    selectedGroupKey={selectedPresetGroupKey}
                    onDoubleClick={(items) => handleBoxItemInsert(items, null, true)}
                    onClose={() => setIsPresetBoxVisible(false)}
                    setMultiInsertConfig={setMultiInsertConfig}
                />
            )}
            <MappingModal 
                isOpen={mappingModal.open}
                onClose={() => setMappingModal({ open: false, itemIndex: null })}
                settings={settings}
                onUpdate={(newSettings) => {
                    setSettings(newSettings);
                    // 如果设置更新了，尝试重新解析当前项
                    if (mappingModal.itemIndex !== null) {
                        handleAutoParseUnprocessed(mappingModal.itemIndex, newSettings);
                    }
                }}
                itemInfo={
                    mappingModal.itemIndex !== null && unprocessedItems[mappingModal.itemIndex] 
                    ? { 
                        productCode: unprocessedItems[mappingModal.itemIndex].details?.productCode || '',
                        notes: unprocessedItems[mappingModal.itemIndex].details?.notes || ''
                    } 
                    : null
                }
            />
            {isLibOpen && (
                <div style={{ position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', backgroundColor: 'rgba(0,0,0,0.5)', zIndex: 3000, display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
                    <div style={{ backgroundColor: 'white', width: '90%', maxWidth: '800px', maxHeight: '90vh', borderRadius: '12px', overflow: 'hidden', display: 'flex', flexDirection: 'column', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1)' }}>
                        <div style={{ padding: '1.5rem', borderBottom: '1px solid #eee', display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#f9fafb' }}>
                            <h2 style={{ margin: 0, fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>📋 策略逻辑库 (Nesting Strategy Library)</h2>
                            <button onClick={() => setIsLibOpen(false)} style={{ border: 'none', background: 'none', fontSize: '1.5rem', cursor: 'pointer', color: '#6b7280' }}>&times;</button>
                        </div>
                        <div style={{ padding: '1rem', borderBottom: '1px solid #eee', backgroundColor: '#fff', display: 'flex', gap: '1rem', alignItems: 'center' }}>
                            <button onClick={exportLibrary} className="button" style={{ fontSize: '0.875rem', padding: '6px 12px', backgroundColor: '#f3f4f6', color: '#374151' }}>📤 导出数据</button>
                            <div style={{ position: 'relative' }}>
                                <button className="button" style={{ fontSize: '0.875rem', padding: '6px 12px', backgroundColor: '#f3f4f6', color: '#374151' }}>📥 导入数据</button>
                                <input type="file" accept=".json" onChange={importLibrary} style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', opacity: 0, cursor: 'pointer' }} />
                            </div>
                            <button onClick={clearLibrary} className="button" style={{ fontSize: '0.875rem', padding: '6px 12px', backgroundColor: '#fee2e2', color: '#b91c1c' }}>🗑️ 清空所有</button>
                            <div style={{ flex: 1 }}></div>
                            <span style={{ fontSize: '0.875rem', color: '#6b7280' }}>* 自动学习利用率 ≥ 90% 的方案</span>
                        </div>
                        <div style={{ flex: 1, overflowY: 'auto', padding: '1.5rem' }}>
                            {strategyLibrary.length === 0 ? (
                                <div style={{ textAlign: 'center', padding: '3rem', color: '#9ca3af' }}>暂无策略，排版成功后系统将自动学习。</div>
                            ) : (
                                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                    <thead style={{ backgroundColor: '#f3f4f6' }}>
                                        <tr>
                                            <th style={{ textAlign: 'left', padding: '12px' }}>名称</th>
                                            <th style={{ textAlign: 'left', padding: '12px' }}>核心参数</th>
                                            <th style={{ textAlign: 'center', padding: '12px' }}>预估利用率</th>
                                            <th style={{ textAlign: 'center', padding: '12px' }}>操作</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {strategyLibrary.map(s => (
                                            <tr key={s.id} style={{ borderBottom: '1px solid #f3f4f6' }}>
                                                <td style={{ padding: '12px' }}>
                                                    <div style={{ fontWeight: '600' }}>{s.name}</div>
                                                    <div style={{ fontSize: '0.75rem', color: '#9ca3af' }}>{new Date(s.createdAt).toLocaleString()}</div>
                                                </td>
                                                <td style={{ padding: '12px' }}>
                                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
                                                        <span style={{ backgroundColor: '#e0e7ff', color: '#4338ca', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem' }}>{s.algorithm}</span>
                                                        <span style={{ backgroundColor: '#fef3c7', color: '#92400e', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem' }}>{s.sortStrategy}</span>
                                                        <span style={{ backgroundColor: '#d1fae5', color: '#065f46', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem' }}>{s.heuristic}</span>
                                                        {s.allowRotation && <span style={{ backgroundColor: '#f3f4f6', color: '#374151', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem' }}>旋转开</span>}
                                                    </div>
                                                </td>
                                                <td style={{ padding: '12px', textAlign: 'center' }}>
                                                    <div style={{ width: '60px', height: '6px', backgroundColor: '#e5e7eb', borderRadius: '3px', margin: '0 auto 4px' }}>
                                                        <div style={{ width: `${s.utilization * 100}%`, height: '100%', backgroundColor: s.utilization > 0.95 ? '#10b981' : '#3b82f6', borderRadius: '3px' }}></div>
                                                    </div>
                                                    <span style={{ fontSize: '0.875rem', fontWeight: 'bold' }}>{Math.round(s.utilization * 100)}%</span>
                                                </td>
                                                <td style={{ padding: '12px', textAlign: 'center' }}>
                                                    <div style={{ display: 'flex', gap: '8px', justifyContent: 'center' }}>
                                                        <button 
                                                            onClick={() => {
                                                                setNestingAlgorithm(s.algorithm);
                                                                setAllowRotation(s.allowRotation);
                                                                setIsLibOpen(false);
                                                                runNesting(s.sortStrategy);
                                                            }} 
                                                            className="button button-primary" 
                                                            style={{ fontSize: '0.75rem', padding: '4px 8px' }}
                                                        >
                                                            应用
                                                        </button>
                                                        <button onClick={() => deleteStrategy(s.id)} style={{ color: '#ef4444', background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.25rem' }}>&times;</button>
                                                    </div>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            )}
                        </div>
                        <div style={{ padding: '1rem', borderTop: '1px solid #eee', textAlign: 'right', backgroundColor: '#f9fafb' }}>
                            <button onClick={() => setIsLibOpen(false)} className="button button-primary" style={{ padding: '8px 24px' }}>关闭系统</button>
                        </div>
                    </div>
                </div>
            )}
        </>
    );
};

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<React.StrictMode><App /></React.StrictMode>);