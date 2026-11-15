// Preset shape definitions (a:prstGeom), converted from ECMA-376 Part 1 presetShapeDefinitions.xml
// (the copy shipped by Apache POI, Apache-2.0) by a one-off script, keeping the common shapes.
// Text rectangles, connection sites and adjust handles are dropped, as are guides the paths do
// not use, which keeps this table small.
//
// Format of each entry: `avLst|gdLst|paths`.
//   avLst: `name default;...` (adjust values, overridable by a:avLst in the document);
//   gdLst: `name op arg arg arg;...` (DrawingML guide formulas, evaluated in order);
//   paths: separated by `!`; each starts with `w h fill stroke` (w/h 0 = shape coordinates;
//          fill N normal, n none, d darken, D darkenLess, l lighten, L lightenLess; stroke 1/0),
//          then commands `M x y`, `L x y`, `A wR hR stAng swAng`, `Q x1 y1 x y`,
//          `C x1 y1 x2 y2 x y`, `Z`.
// Shapes not listed render as their bounding rectangle (with a warning).

/** Compact preset geometry table (see the file header for the format). */
export const PRESETS: Readonly<Record<string, string>> = {
  rect: '||0 0 N 1 M l t L r t L r b L l b Z',
  roundRect:
    'adj 16667|a pin 0 adj 50000;x1 */ ss a 100000;x2 +- r 0 x1;y2 +- b 0 x1|0 0 N 1 M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z',
  ellipse:
    '||0 0 N 1 M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  triangle: 'adj 50000|a pin 0 adj 100000;x2 */ w a 100000|0 0 N 1 M l b L x2 t L r b Z',
  rtTriangle: '||0 0 N 1 M l b L l t L r b Z',
  diamond: '||0 0 N 1 M l vc L hc t L r vc L hc b Z',
  parallelogram:
    'adj 25000|maxAdj */ 100000 w ss;a pin 0 adj maxAdj;x2 */ ss a 100000;x5 +- r 0 x2|0 0 N 1 M l b L x2 t L r t L x5 b Z',
  trapezoid:
    'adj 25000|maxAdj */ 50000 w ss;a pin 0 adj maxAdj;x2 */ ss a 100000;x3 +- r 0 x2|0 0 N 1 M l b L x2 t L x3 t L r b Z',
  pentagon:
    'hf 105146;vf 110557|swd2 */ wd2 hf 100000;shd2 */ hd2 vf 100000;svc */ vc vf 100000;dx1 cos swd2 1080000;dx2 cos swd2 18360000;dy1 sin shd2 1080000;dy2 sin shd2 18360000;x1 +- hc 0 dx1;x2 +- hc 0 dx2;x3 +- hc dx2 0;x4 +- hc dx1 0;y1 +- svc 0 dy1;y2 +- svc 0 dy2|0 0 N 1 M x1 y1 L hc t L x4 y1 L x3 y2 L x2 y2 Z',
  hexagon:
    'adj 25000;vf 115470|maxAdj */ 50000 w ss;a pin 0 adj maxAdj;shd2 */ hd2 vf 100000;x1 */ ss a 100000;x2 +- r 0 x1;dy1 sin shd2 3600000;y1 +- vc 0 dy1;y2 +- vc dy1 0|0 0 N 1 M l vc L x1 y1 L x2 y1 L r vc L x2 y2 L x1 y2 Z',
  octagon:
    'adj 29289|a pin 0 adj 50000;x1 */ ss a 100000;x2 +- r 0 x1;y2 +- b 0 x1|0 0 N 1 M l x1 L x1 t L x2 t L r x1 L r y2 L x2 b L x1 b L l y2 Z',
  star4:
    'adj 12500|a pin 0 adj 50000;iwd2 */ wd2 a 50000;ihd2 */ hd2 a 50000;sdx cos iwd2 2700000;sdy sin ihd2 2700000;sx1 +- hc 0 sdx;sx2 +- hc sdx 0;sy1 +- vc 0 sdy;sy2 +- vc sdy 0|0 0 N 1 M l vc L sx1 sy1 L hc t L sx2 sy1 L r vc L sx2 sy2 L hc b L sx1 sy2 Z',
  star5:
    'adj 19098;hf 105146;vf 110557|a pin 0 adj 50000;swd2 */ wd2 hf 100000;shd2 */ hd2 vf 100000;svc */ vc vf 100000;dx1 cos swd2 1080000;dx2 cos swd2 18360000;dy1 sin shd2 1080000;dy2 sin shd2 18360000;x1 +- hc 0 dx1;x2 +- hc 0 dx2;x3 +- hc dx2 0;x4 +- hc dx1 0;y1 +- svc 0 dy1;y2 +- svc 0 dy2;iwd2 */ swd2 a 50000;ihd2 */ shd2 a 50000;sdx1 cos iwd2 20520000;sdx2 cos iwd2 3240000;sdy1 sin ihd2 3240000;sdy2 sin ihd2 20520000;sx1 +- hc 0 sdx1;sx2 +- hc 0 sdx2;sx3 +- hc sdx2 0;sx4 +- hc sdx1 0;sy1 +- svc 0 sdy1;sy2 +- svc 0 sdy2;sy3 +- svc ihd2 0|0 0 N 1 M x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x4 y1 L sx4 sy2 L x3 y2 L hc sy3 L x2 y2 L sx1 sy2 Z',
  star6:
    'adj 28868;hf 115470|a pin 0 adj 50000;swd2 */ wd2 hf 100000;dx1 cos swd2 1800000;x1 +- hc 0 dx1;x2 +- hc dx1 0;y2 +- vc hd4 0;iwd2 */ swd2 a 50000;ihd2 */ hd2 a 50000;sdx2 */ iwd2 1 2;sx1 +- hc 0 iwd2;sx2 +- hc 0 sdx2;sx3 +- hc sdx2 0;sx4 +- hc iwd2 0;sdy1 sin ihd2 3600000;sy1 +- vc 0 sdy1;sy2 +- vc sdy1 0|0 0 N 1 M x1 hd4 L sx2 sy1 L hc t L sx3 sy1 L x2 hd4 L sx4 vc L x2 y2 L sx3 sy2 L hc b L sx2 sy2 L x1 y2 L sx1 vc Z',
  star8:
    'adj 37500|a pin 0 adj 50000;dx1 cos wd2 2700000;x1 +- hc 0 dx1;x2 +- hc dx1 0;dy1 sin hd2 2700000;y1 +- vc 0 dy1;y2 +- vc dy1 0;iwd2 */ wd2 a 50000;ihd2 */ hd2 a 50000;sdx1 */ iwd2 92388 100000;sdx2 */ iwd2 38268 100000;sdy1 */ ihd2 92388 100000;sdy2 */ ihd2 38268 100000;sx1 +- hc 0 sdx1;sx2 +- hc 0 sdx2;sx3 +- hc sdx2 0;sx4 +- hc sdx1 0;sy1 +- vc 0 sdy1;sy2 +- vc 0 sdy2;sy3 +- vc sdy2 0;sy4 +- vc sdy1 0|0 0 N 1 M l vc L sx1 sy2 L x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x2 y1 L sx4 sy2 L r vc L sx4 sy3 L x2 y2 L sx3 sy4 L hc b L sx2 sy4 L x1 y2 L sx1 sy3 Z',
  rightArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 100000 w ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;dx1 */ ss a2 100000;x1 +- r 0 dx1;dy1 */ h a1 200000;y1 +- vc 0 dy1;y2 +- vc dy1 0|0 0 N 1 M l y1 L x1 y1 L x1 t L r vc L x1 b L x1 y2 L l y2 Z',
  leftArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 100000 w ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;dx2 */ ss a2 100000;x2 +- l dx2 0;dy1 */ h a1 200000;y1 +- vc 0 dy1;y2 +- vc dy1 0|0 0 N 1 M l vc L x2 t L x2 y1 L r y1 L r y2 L x2 y2 L x2 b Z',
  upArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 100000 h ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;dy2 */ ss a2 100000;y2 +- t dy2 0;dx1 */ w a1 200000;x1 +- hc 0 dx1;x2 +- hc dx1 0|0 0 N 1 M l y2 L hc t L r y2 L x2 y2 L x2 b L x1 b L x1 y2 Z',
  downArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 100000 h ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;dy1 */ ss a2 100000;y1 +- b 0 dy1;dx1 */ w a1 200000;x1 +- hc 0 dx1;x2 +- hc dx1 0|0 0 N 1 M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L hc b Z',
  leftRightArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 50000 w ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;x2 */ ss a2 100000;x3 +- r 0 x2;dy */ h a1 200000;y1 +- vc 0 dy;y2 +- vc dy 0|0 0 N 1 M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z',
  upDownArrow:
    'adj1 50000;adj2 50000|maxAdj2 */ 50000 h ss;a1 pin 0 adj1 100000;a2 pin 0 adj2 maxAdj2;y2 */ ss a2 100000;y3 +- b 0 y2;dx1 */ w a1 200000;x1 +- hc 0 dx1;x2 +- hc dx1 0|0 0 N 1 M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z',
  chevron:
    'adj 50000|maxAdj */ 100000 w ss;a pin 0 adj maxAdj;x1 */ ss a 100000;x2 +- r 0 x1|0 0 N 1 M l t L x2 t L r vc L x2 b L l b L x1 vc Z',
  homePlate:
    'adj 50000|maxAdj */ 100000 w ss;a pin 0 adj maxAdj;dx1 */ ss a 100000;x1 +- r 0 dx1|0 0 N 1 M l t L x1 t L r vc L x1 b L l b Z',
  plus: 'adj 25000|a pin 0 adj 50000;x1 */ ss a 100000;x2 +- r 0 x1;y2 +- b 0 x1|0 0 N 1 M l x1 L x1 x1 L x1 t L x2 t L x2 x1 L r x1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 Z',
  line: '||0 0 N 1 M l t L r b',
  straightConnector1: '||0 0 n 1 M l t L r b',
  bentConnector2: '||0 0 n 1 M l t L r t L r b',
  bentConnector3: 'adj1 50000|x1 */ w adj1 100000|0 0 n 1 M l t L x1 t L x1 b L r b',
  curvedConnector2: '||0 0 n 1 M l t C wd2 t r hd2 r b',
  curvedConnector3:
    'adj1 50000|x2 */ w adj1 100000;x1 +/ l x2 2;x3 +/ r x2 2;y3 */ h 3 4|0 0 n 1 M l t C x1 t x2 hd4 x2 vc C x2 y3 x3 b r b',
  wedgeRectCallout:
    'adj1 -20833;adj2 62500|dxPos */ w adj1 100000;dyPos */ h adj2 100000;xPos +- hc dxPos 0;yPos +- vc dyPos 0;dq */ dxPos h w;ady abs dyPos;adq abs dq;dz +- ady 0 adq;xg1 ?: dxPos 7 2;xg2 ?: dxPos 10 5;x1 */ w xg1 12;x2 */ w xg2 12;yg1 ?: dyPos 7 2;yg2 ?: dyPos 10 5;y1 */ h yg1 12;y2 */ h yg2 12;t1 ?: dxPos l xPos;xl ?: dz l t1;t2 ?: dyPos x1 xPos;xt ?: dz t2 x1;t3 ?: dxPos xPos r;xr ?: dz r t3;t4 ?: dyPos xPos x1;xb ?: dz t4 x1;t5 ?: dxPos y1 yPos;yl ?: dz y1 t5;t6 ?: dyPos t yPos;yt ?: dz t6 t;t7 ?: dxPos yPos y1;yr ?: dz y1 t7;t8 ?: dyPos yPos b;yb ?: dz t8 b|0 0 N 1 M l t L x1 t L xt yt L x2 t L r t L r y1 L xr yr L r y2 L r b L x2 b L xb yb L x1 b L l b L l y2 L xl yl L l y1 Z',
  wedgeRoundRectCallout:
    'adj1 -20833;adj2 62500;adj3 16667|dxPos */ w adj1 100000;dyPos */ h adj2 100000;xPos +- hc dxPos 0;yPos +- vc dyPos 0;dq */ dxPos h w;ady abs dyPos;adq abs dq;dz +- ady 0 adq;xg1 ?: dxPos 7 2;xg2 ?: dxPos 10 5;x1 */ w xg1 12;x2 */ w xg2 12;yg1 ?: dyPos 7 2;yg2 ?: dyPos 10 5;y1 */ h yg1 12;y2 */ h yg2 12;t1 ?: dxPos l xPos;xl ?: dz l t1;t2 ?: dyPos x1 xPos;xt ?: dz t2 x1;t3 ?: dxPos xPos r;xr ?: dz r t3;t4 ?: dyPos xPos x1;xb ?: dz t4 x1;t5 ?: dxPos y1 yPos;yl ?: dz y1 t5;t6 ?: dyPos t yPos;yt ?: dz t6 t;t7 ?: dxPos yPos y1;yr ?: dz y1 t7;t8 ?: dyPos yPos b;yb ?: dz t8 b;u1 */ ss adj3 100000;u2 +- r 0 u1;v2 +- b 0 u1|0 0 N 1 M l u1 A u1 u1 cd2 cd4 L x1 t L xt yt L x2 t L u2 t A u1 u1 3cd4 cd4 L r y1 L xr yr L r y2 L r v2 A u1 u1 0 cd4 L x2 b L xb yb L x1 b L u1 b A u1 u1 cd4 cd4 L l y2 L xl yl L l y1 Z',
  wedgeEllipseCallout:
    'adj1 -20833;adj2 62500|dxPos */ w adj1 100000;dyPos */ h adj2 100000;xPos +- hc dxPos 0;yPos +- vc dyPos 0;sdx */ dxPos h 1;sdy */ dyPos w 1;pang at2 sdx sdy;stAng +- pang 660000 0;enAng +- pang 0 660000;dx1 cos wd2 stAng;dy1 sin hd2 stAng;x1 +- hc dx1 0;y1 +- vc dy1 0;dx2 cos wd2 enAng;dy2 sin hd2 enAng;stAng1 at2 dx1 dy1;enAng1 at2 dx2 dy2;swAng1 +- enAng1 0 stAng1;swAng2 +- swAng1 21600000 0;swAng ?: swAng1 swAng1 swAng2|0 0 N 1 M xPos yPos L x1 y1 A wd2 hd2 stAng1 swAng Z',
  cloud:
    '||43200 43200 N 1 M 3900 14370 A 6753 9190 -11429249 7426832 A 5333 7267 -8646143 5396714 A 4365 5945 -8748475 5983381 A 4857 6595 -7859164 7034504 A 5333 7273 -4722533 6541615 A 6775 9220 -2776035 7816140 A 5785 7867 37501 6842000 A 6752 9215 1347096 6910353 A 7720 10543 3974558 4542661 A 4360 5918 -16496525 8804134 A 4345 5945 -14809710 9151131 Z!43200 43200 n 1 M 4693 26177 A 4345 5945 5204520 1585770 M 6928 34899 A 4360 5918 4416628 686848 M 16478 39090 A 6752 9215 8257449 844866 M 28827 34751 A 6752 9215 387196 959901 M 34129 22954 A 5785 7867 -4217541 4255042 M 41798 15354 A 5333 7273 1819082 1665090 M 38324 5426 A 4857 6595 -824660 891534 M 29078 3952 A 4857 6595 -8950887 1091722 M 22141 4720 A 4365 5945 -9809656 1061181 M 14000 5192 A 6753 9190 -4002417 739161 M 4127 15789 A 6753 9190 9459261 711490',
  heart:
    '|dx1 */ w 49 48;dx2 */ w 10 48;x1 +- hc 0 dx1;x2 +- hc 0 dx2;x3 +- hc dx2 0;x4 +- hc dx1 0;y1 +- t 0 hd3|0 0 N 1 M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z',
  flowChartProcess: '||1 1 N 1 M 0 0 L 1 0 L 1 1 L 0 1 Z',
  flowChartAlternateProcess:
    '|x2 +- r 0 ssd6;y2 +- b 0 ssd6|0 0 N 1 M l ssd6 A ssd6 ssd6 cd2 cd4 L x2 t A ssd6 ssd6 3cd4 cd4 L r y2 A ssd6 ssd6 0 cd4 L ssd6 b A ssd6 ssd6 cd4 cd4 Z',
  flowChartDecision: '||2 2 N 1 M 0 1 L 1 0 L 2 1 L 1 2 Z',
  flowChartTerminator:
    '||21600 21600 N 1 M 3475 0 L 18125 0 A 3475 10800 3cd4 cd2 L 3475 21600 A 3475 10800 cd4 cd2 Z',
  flowChartConnector:
    '||0 0 N 1 M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  flowChartDocument:
    '||21600 21600 N 1 M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z',
  flowChartInputOutput: '||5 5 N 1 M 0 5 L 1 0 L 5 0 L 4 5 Z',
  flowChartPredefinedProcess:
    '||1 1 N 0 M 0 0 L 1 0 L 1 1 L 0 1 Z!8 8 n 1 M 1 0 L 1 8 M 7 0 L 7 8!1 1 n 1 M 0 0 L 1 0 L 1 1 L 0 1 Z',
  can: 'adj 25000|maxAdj */ 50000 h ss;a pin 0 adj maxAdj;y1 */ ss a 200000;y3 +- b 0 y1|0 0 N 0 M l y1 A wd2 y1 cd2 -10800000 L r y3 A wd2 y1 0 cd2 Z!0 0 l 0 M l y1 A wd2 y1 cd2 cd2 A wd2 y1 0 cd2 Z!0 0 n 1 M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 L l y1',
  cube: 'adj 25000|a pin 0 adj 100000;y1 */ ss a 100000;y4 +- b 0 y1;x4 +- r 0 y1|0 0 N 0 M l y1 L x4 y1 L x4 b L l b Z!0 0 D 0 M x4 y1 L r t L r y4 L x4 b Z!0 0 L 0 M l y1 L y1 t L r t L x4 y1 Z!0 0 n 1 M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b',
  frame:
    'adj1 12500|a1 pin 0 adj1 50000;x1 */ ss a1 100000;x4 +- r 0 x1;y4 +- b 0 x1|0 0 N 1 M l t L r t L r b L l b Z M x1 x1 L x1 y4 L x4 y4 L x4 x1 Z',
  donut:
    'adj 25000|a pin 0 adj 50000;dr */ ss a 100000;iwd2 +- wd2 0 dr;ihd2 +- hd2 0 dr|0 0 N 1 M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z',
  blockArc:
    'adj1 10800000;adj2 0;adj3 25000|stAng pin 0 adj1 21599999;istAng pin 0 adj2 21599999;a3 pin 0 adj3 50000;sw11 +- istAng 0 stAng;sw12 +- sw11 21600000 0;swAng ?: sw11 sw11 sw12;iswAng +- 0 0 swAng;wt1 sin wd2 stAng;ht1 cos hd2 stAng;dx1 cat2 wd2 ht1 wt1;dy1 sat2 hd2 ht1 wt1;x1 +- hc dx1 0;y1 +- vc dy1 0;dr */ ss a3 100000;iwd2 +- wd2 0 dr;ihd2 +- hd2 0 dr;wt2 sin iwd2 istAng;ht2 cos ihd2 istAng;dx2 cat2 iwd2 ht2 wt2;dy2 sat2 ihd2 ht2 wt2;x2 +- hc dx2 0;y2 +- vc dy2 0|0 0 N 1 M x1 y1 A wd2 hd2 stAng swAng L x2 y2 A iwd2 ihd2 istAng iswAng Z',
  arc: 'adj1 16200000;adj2 0|stAng pin 0 adj1 21599999;enAng pin 0 adj2 21599999;sw11 +- enAng 0 stAng;sw12 +- sw11 21600000 0;swAng ?: sw11 sw11 sw12;wt1 sin wd2 stAng;ht1 cos hd2 stAng;dx1 cat2 wd2 ht1 wt1;dy1 sat2 hd2 ht1 wt1;x1 +- hc dx1 0;y1 +- vc dy1 0|0 0 N 0 M x1 y1 A wd2 hd2 stAng swAng L hc vc Z!0 0 n 1 M x1 y1 A wd2 hd2 stAng swAng',
  bracketPair:
    'adj 16667|a pin 0 adj 50000;x1 */ ss a 100000;x2 +- r 0 x1;y2 +- b 0 x1|0 0 N 0 M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z!0 0 n 1 M x1 b A x1 x1 cd4 cd4 L l x1 A x1 x1 cd2 cd4 M x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4',
  bracePair:
    'adj 8333|a pin 0 adj 25000;x1 */ ss a 100000;x2 */ ss a 50000;x3 +- r 0 x2;x4 +- r 0 x1;y2 +- vc 0 x1;y3 +- vc x1 0;y4 +- b 0 x1|0 0 N 0 M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 L x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4 Z!0 0 n 1 M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 M x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4',
  leftBracket:
    'adj 8333|maxAdj */ 50000 h ss;a pin 0 adj maxAdj;y1 */ ss a 100000|0 0 N 0 M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4 Z!0 0 n 1 M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4',
  rightBracket:
    'adj 8333|maxAdj */ 50000 h ss;a pin 0 adj maxAdj;y1 */ ss a 100000;y2 +- b 0 y1|0 0 N 0 M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4 Z!0 0 n 1 M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4',
  leftBrace:
    'adj1 8333;adj2 50000|a2 pin 0 adj2 100000;q1 +- 100000 0 a2;q2 min q1 a2;q3 */ q2 1 2;maxAdj1 */ q3 h ss;a1 pin 0 adj1 maxAdj1;y1 */ ss a1 100000;y3 */ h a2 100000;y4 +- y3 y1 0|0 0 N 0 M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4 Z!0 0 n 1 M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4',
  rightBrace:
    'adj1 8333;adj2 50000|a2 pin 0 adj2 100000;q1 +- 100000 0 a2;q2 min q1 a2;q3 */ q2 1 2;maxAdj1 */ q3 h ss;a1 pin 0 adj1 maxAdj1;y1 */ ss a1 100000;y3 */ h a2 100000;y2 +- y3 0 y1;y4 +- b 0 y1|0 0 N 0 M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4 Z!0 0 n 1 M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4',
  snip1Rect:
    'adj 16667|a pin 0 adj 50000;dx1 */ ss a 100000;x1 +- r 0 dx1|0 0 N 1 M l t L x1 t L r dx1 L r b L l b Z',
  snip2SameRect:
    'adj1 16667;adj2 0|a1 pin 0 adj1 50000;a2 pin 0 adj2 50000;tx1 */ ss a1 100000;tx2 +- r 0 tx1;bx1 */ ss a2 100000;bx2 +- r 0 bx1;by1 +- b 0 bx1|0 0 N 1 M tx1 t L tx2 t L r tx1 L r by1 L bx2 b L bx1 b L l by1 L l tx1 Z',
  round1Rect:
    'adj 16667|a pin 0 adj 50000;dx1 */ ss a 100000;x1 +- r 0 dx1|0 0 N 1 M l t L x1 t A dx1 dx1 3cd4 cd4 L r b L l b Z',
  round2SameRect:
    'adj1 16667;adj2 0|a1 pin 0 adj1 50000;a2 pin 0 adj2 50000;tx1 */ ss a1 100000;tx2 +- r 0 tx1;bx1 */ ss a2 100000;by1 +- b 0 bx1|0 0 N 1 M tx1 t L tx2 t A tx1 tx1 3cd4 cd4 L r by1 A bx1 bx1 0 cd4 L bx1 b A bx1 bx1 cd4 cd4 L l tx1 A tx1 tx1 cd2 cd4 Z',
  round2DiagRect:
    'adj1 16667;adj2 0|a1 pin 0 adj1 50000;a2 pin 0 adj2 50000;x1 */ ss a1 100000;y1 +- b 0 x1;a */ ss a2 100000;x2 +- r 0 a|0 0 N 1 M x1 t L x2 t A a a 3cd4 cd4 L r y1 A x1 x1 0 cd4 L a b A a a cd4 cd4 L l x1 A x1 x1 cd2 cd4 Z',
  teardrop:
    'adj 100000|a pin 0 adj 200000;r2 sqrt 2;tw */ wd2 r2 1;th */ hd2 r2 1;sw */ tw a 100000;sh */ th a 100000;dx1 cos sw 2700000;dy1 sin sh 2700000;x1 +- hc dx1 0;y1 +- vc 0 dy1;x2 +/ hc x1 2;y2 +/ vc y1 2|0 0 N 1 M l vc A wd2 hd2 cd2 cd4 Q x2 t x1 y1 Q r y2 r vc A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  smileyFace:
    'adj 4653|a pin -4653 adj 4653;x1 */ w 4969 21699;x2 */ w 6215 21600;x3 */ w 13135 21600;x4 */ w 16640 21600;y1 */ h 7570 21600;y3 */ h 16515 21600;dy2 */ h a 100000;y2 +- y3 0 dy2;y4 +- y3 dy2 0;dy3 */ h a 50000;y5 +- y4 dy3 0;wR */ w 1125 21600;hR */ h 1125 21600|0 0 N 0 M l vc A wd2 hd2 cd2 21600000 Z!0 0 D 1 M x2 y1 A wR hR cd2 21600000 M x3 y1 A wR hR cd2 21600000!0 0 n 1 M x1 y2 Q hc y5 x4 y2!0 0 n 1 M l vc A wd2 hd2 cd2 21600000 Z',
};
