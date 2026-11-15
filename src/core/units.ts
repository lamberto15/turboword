// Unit conversions. WordprocessingML measures most lengths in twentieths of a point (twips),
// font sizes in half-points and DrawingML extents in EMUs (English Metric Units).

/** EMUs per point. */
export const EMU_PER_PT = 12_700;
/** EMUs per inch. */
export const EMU_PER_INCH = 914_400;
/** Twips per point. */
export const TWIPS_PER_PT = 20;
/** Twips per inch. */
export const TWIPS_PER_INCH = 1440;

export const twipsToPt = (twips: number): number => twips / TWIPS_PER_PT;
export const ptToTwips = (pt: number): number => Math.round(pt * TWIPS_PER_PT);
export const emuToPt = (emu: number): number => emu / EMU_PER_PT;
export const ptToEmu = (pt: number): number => Math.round(pt * EMU_PER_PT);
