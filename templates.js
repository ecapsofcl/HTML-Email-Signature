/**
 * eCAPS signature renderer (shared).
 * Keep these two copies IDENTICAL:
 *   1 - Upload to GitHub/templates.js          (portal live preview)
 *   2 - Paste into Google Apps Script/Templates.gs (files that get emailed)
 *
 * Email-client rules followed everywhere below:
 *   - tables only (no div/float/flex), every cell has explicit padding
 *   - every image has width AND height attributes (Outlook resizes otherwise)
 *   - px font sizes + mso-line-height-rule:exactly so line spacing matches in Outlook
 *   - contact details in a 2-column table so labels and values line up in every client
 */
var SigTemplates = (function () {
  var DEFAULT_FONT = 'Helvetica, Arial, sans-serif';
  var ASSETS = 'https://s3.amazonaws.com/htmlsig-assets/';

  // Set at the start of every render() from Style settings
  var FONT = 'font-family:' + DEFAULT_FONT + ';';
  var INK = '#212121', MUTE = '#6b6b6b', LINK = '#477ccc';
  var LABELMODE = '', SEPMODE = '', ICONDATA = null;

  var TABS = [['main', 'Details'], ['social', 'Social links'], ['icons', 'Icon design'], ['style', 'Layout & style'],
              ['banner', 'Banner'], ['extras', 'Extras'], ['badges', 'Badges'], ['disclaimer', 'Disclaimer']];

  var FONTS = [
    ['Helvetica, Arial, sans-serif', 'Sans Serif (Arial / Helvetica)'], ['Georgia, serif', 'Serif (Georgia)'],
    ['Arial, Helvetica, sans-serif', 'Arial'], ["'Arial Black', Gadget, sans-serif", 'Arial Black'],
    ["'Arial Narrow', Arial, sans-serif", 'Arial Narrow'], ["'Arial Rounded MT Bold', Arial, sans-serif", 'Arial Rounded MT Bold'],
    ['Calibri, Candara, Segoe, sans-serif', 'Calibri'], ['Candara, Calibri, Segoe, sans-serif', 'Candara'],
    ["'Century Gothic', AppleGothic, sans-serif", 'Century Gothic'], ["'Gill Sans', 'Gill Sans MT', Calibri, sans-serif", 'Gill Sans'],
    ["'Segoe UI', Tahoma, sans-serif", 'Segoe UI'], ['Tahoma, Geneva, sans-serif', 'Tahoma'], ['Verdana, Geneva, sans-serif', 'Verdana'],
    ["'Trebuchet MS', Helvetica, sans-serif", 'Trebuchet MS'], ["'Book Antiqua', Palatino, serif", 'Book Antiqua'],
    ['Cambria, Georgia, serif', 'Cambria'], ['Garamond, Baskerville, serif', 'Garamond'],
    ["'Lucida Bright', Georgia, serif", 'Lucida Bright'], ["Palatino, 'Palatino Linotype', serif", 'Palatino'],
    ["Baskerville, 'Times New Roman', serif", 'Baskerville'], ["'Times New Roman', Times, serif", 'Times New Roman'],
    ["'Courier New', Courier, monospace", 'Courier New'], ["'Lucida Sans Typewriter', 'Lucida Console', monospace", 'Lucida Sans Typewriter']
  ];

  /* Social networks: [key, label, example link, current brand colour, (unused), symbol colour on brand background]
     Icon artwork lives in icons.js (portal only). */
  var SOCIAL = [
    ['twitter', 'X (Twitter)', 'https://x.com/yourcompany', '#000000'],
    ['facebook', 'Facebook', 'https://www.facebook.com/yourcompany', '#0866FF'],
    ['linkedin', 'LinkedIn', 'https://www.linkedin.com/company/yourcompany', '#0A66C2'],
    ['instagram', 'Instagram', 'https://www.instagram.com/yourcompany', '#FF0069'],
    ['youtube', 'YouTube', 'https://www.youtube.com/@yourcompany', '#FF0000'],
    ['whatsapp', 'WhatsApp', 'https://wa.me/919876543210', '#25D366'],
    ['teams', 'Microsoft Teams', 'https://teams.microsoft.com/l/chat/0/0?users=name@caps.in', '#5059C9'],
    ['skype', 'Skype', 'skype:yourname?chat', '#00AFF0'],
    ['telegram', 'Telegram', 'https://t.me/yourname', '#26A5E4'],
    ['wechat', 'WeChat', 'weixin://dl/chat?yourid', '#07C160'],
    ['tiktok', 'TikTok', 'https://www.tiktok.com/@yourcompany', '#000000'],
    ['pinterest', 'Pinterest', 'https://www.pinterest.com/yourcompany', '#BD081C'],
    ['maps', 'Google Maps', 'https://maps.app.goo.gl/yourplace', '#4285F4'],
    ['github', 'GitHub', 'https://github.com/yourcompany', '#181717'],
    ['yelp', 'Yelp', 'https://www.yelp.com/biz/yourcompany', '#FF1A1A'],
    ['tripadvisor', 'TripAdvisor', 'https://www.tripadvisor.com/yourlisting', '#34E0A1', '', '#000000'],
    ['snapchat', 'Snapchat', 'https://www.snapchat.com/add/yourname', '#FFFC00', '', '#000000'],
    ['spotify', 'Spotify', 'https://open.spotify.com/user/yourname', '#1ED760', '', '#000000'],
    ['xing', 'Xing', 'https://www.xing.com/profile/yourname', '#006567'],
    ['substack', 'Substack', 'https://yourname.substack.com', '#FF6719'],
    ['behance', 'Behance', 'https://www.behance.net/yourname', '#1769FF'],
    ['dribbble', 'Dribbble', 'https://dribbble.com/yourname', '#EA4C89'],
    ['vimeo', 'Vimeo', 'https://vimeo.com/yourcompany', '#1AB7EA'],
    ['flickr', 'Flickr', 'https://www.flickr.com/photos/yourname', '#0063DC'],
    ['fivehundredpx', '500px', 'https://500px.com/p/yourname', '#222222'],
    ['soundcloud', 'SoundCloud', 'https://soundcloud.com/yourname', '#FF5500'],
    ['mixcloud', 'Mixcloud', 'https://www.mixcloud.com/yourname', '#5000FF'],
    ['tumblr', 'Tumblr', 'https://yourname.tumblr.com', '#36465D'],
    ['wordpress', 'WordPress', 'https://yourblog.wordpress.com', '#21759B'],
    ['quora', 'Quora', 'https://www.quora.com/profile/yourname', '#B92B27'],
    ['stackoverflow', 'Stack Overflow', 'https://stackoverflow.com/users/000000/yourname', '#F58025'],
    ['bitbucket', 'Bitbucket', 'https://bitbucket.org/yourcompany', '#0052CC'],
    ['imdb', 'IMDb', 'https://www.imdb.com/name/nm0000000', '#F5C518', '', '#000000'],
    ['zillow', 'Zillow', 'https://www.zillow.com/profile/yourname', '#006AFF'],
    ['houzz', 'Houzz', 'https://www.houzz.com/pro/yourname', '#4DBC15'],
    ['periscope', 'Periscope', 'https://www.pscp.tv/yourname', '#40A4C4'],
    ['googleplus', 'Google+', 'https://plus.google.com/yourpage', '#DB4437']
  ];

  /* Company-level fields: general value in Settings, can be overridden per branch or person.
     type: text | url | long | color | select | number;  upload: image can be uploaded */
  var FIELDS = [
    { key: 'company_name',     label: 'Company name',            tab: 'main' },
    { key: 'website',          label: 'Website',                 tab: 'main', type: 'url', example: 'https://www.caps.in' },
    { key: 'office_phone',     label: 'Office phone',            tab: 'main' },
    { key: 'fax',              label: 'Fax',                     tab: 'main' },
    { key: 'address',          label: 'Address',                 tab: 'main' },
    { key: 'address2',         label: 'Address line 2',          tab: 'main' },
    { key: 'office_locations', label: 'Office locations',        tab: 'main', wide: true, hint: 'Comma separated' },
    { key: 'logo_url',         label: 'Logo',                    tab: 'main', type: 'url', upload: true, wide: true },

    { key: 'disclaimer',       label: 'Disclaimer text',         tab: 'disclaimer', type: 'long', wide: true },

    { key: 'banner_url',       label: 'Banner image',            tab: 'banner', type: 'url', upload: true, wide: true, hint: 'PNG, JPG or GIF up to 2 MB. For a sharp banner, upload it at twice the display width.' },
    { key: 'banner_link',      label: 'Banner link',             tab: 'banner', type: 'url', example: 'https://store.caps.in' },
    { key: 'banner_width',     label: 'Display width (px)',      tab: 'banner', type: 'number', example: '450' },

    { key: 'app_apple',        label: 'Apple App Store link',    tab: 'extras', type: 'url', example: 'https://apps.apple.com/app/id000000000' },
    { key: 'app_google',       label: 'Google Play link',        tab: 'extras', type: 'url', example: 'https://play.google.com/store/apps/details?id=com.yourapp' },
    { key: 'app_amazon',       label: 'Amazon Appstore link',    tab: 'extras', type: 'url', example: 'https://www.amazon.com/dp/B000000000' },
    { key: 'calendar_link',    label: 'Meeting booking link',    tab: 'extras', type: 'url', example: 'https://calendly.com/yourname', hint: 'Adds a "Schedule a meeting" button' },
    { key: 'whatsapp_number',  label: 'WhatsApp chat number',    tab: 'extras', example: '919876543210', hint: 'Adds a "Contact me on WhatsApp" button' },
    { key: 'whatsapp_message', label: 'WhatsApp opening message', tab: 'extras', example: 'Hi, I would like to know more' },
    { key: 'tagline',          label: 'Tagline or quote',        tab: 'extras', wide: true },
    { key: 'cta_text',         label: 'Button text',             tab: 'extras', example: 'Visit our store' },
    { key: 'cta_link',         label: 'Button link',             tab: 'extras', type: 'url', example: 'https://store.caps.in' },
    { key: 'eco_note',         label: 'Footer note',             tab: 'extras', wide: true, example: 'Please consider the environment before printing this email.' },

    { key: 'font_family',      label: 'Font',                    tab: 'style', type: 'select', options: FONTS },
    { key: 'font_size',        label: 'Text size',               tab: 'style', type: 'select', options: [['small', 'Small'], ['normal', 'Normal'], ['large', 'Large']] },
    { key: 'brand_color',      label: 'Accent colour',           tab: 'style', type: 'color' },
    { key: 'text_color',       label: 'Text colour',             tab: 'style', type: 'color' },
    { key: 'secondary_color',  label: 'Label colour',            tab: 'style', type: 'color' },
    { key: 'link_color',       label: 'Link colour',             tab: 'style', type: 'color' },
    { key: 'separator',        label: 'Separators',              tab: 'style', type: 'select', options: [['/', 'Slash  /'], ['|', 'Bar  |'], ['none', 'None']] },
    { key: 'phone_labels',     label: 'Contact labels',          tab: 'style', type: 'select', options: [['text', 'Words (Mobile:)'], ['letters', 'Letters (M)'], ['none', 'None']] },
    { key: 'logo_width',       label: 'Logo width (px)',         tab: 'style', type: 'number' },

    { key: 'icon_style',       label: 'Icon design',             tab: 'icons', type: 'select', options: [['official', 'Official logos (full colour)'], ['custom', 'Custom (your shape and colours)'], ['round', 'Standard: round'], ['square', 'Standard: square'], ['rounded', 'Standard: rounded'], ['polygon', 'Standard: polygon'], ['grey', 'Standard: grey']] },
    { key: 'icon_shape',       label: 'Shape',                   tab: 'icons', type: 'select', options: [['circle', 'Circle'], ['rounded', 'Rounded square'], ['square', 'Square'], ['plain', 'Icon only (no background)']] },
    { key: 'icon_color_mode',  label: 'Colours',                 tab: 'icons', type: 'select', options: [['single', 'One colour for all'], ['brand', 'Each network\'s own colour']] },
    { key: 'icon_bg',          label: 'Background colour',       tab: 'icons', type: 'color' },
    { key: 'icon_fg',          label: 'Symbol colour',           tab: 'icons', type: 'color' },
    { key: 'icon_size',        label: 'Size',                    tab: 'icons', type: 'select', options: [['16', 'Small (16px)'], ['20', 'Medium (20px)'], ['24', 'Large (24px)'], ['32', 'Extra large (32px)']] },
    { key: 'icon_gap',         label: 'Gap between icons (px)',  tab: 'icons', type: 'number' },
    { key: 'icon_base_url',    label: 'Own icon folder URL',     tab: 'icons', type: 'url', wide: true, hint: 'Advanced. Leave empty unless you host your own icon files.' },

    { key: 'badge1_img',  label: 'Badge 1 image', tab: 'badges', type: 'url', upload: true },
    { key: 'badge1_link', label: 'Badge 1 link',  tab: 'badges', type: 'url' },
    { key: 'badge2_img',  label: 'Badge 2 image', tab: 'badges', type: 'url', upload: true },
    { key: 'badge2_link', label: 'Badge 2 link',  tab: 'badges', type: 'url' },
    { key: 'badge3_img',  label: 'Badge 3 image', tab: 'badges', type: 'url', upload: true },
    { key: 'badge3_link', label: 'Badge 3 link',  tab: 'badges', type: 'url' },
    { key: 'badge4_img',  label: 'Badge 4 image', tab: 'badges', type: 'url', upload: true },
    { key: 'badge4_link', label: 'Badge 4 link',  tab: 'badges', type: 'url' },
    { key: 'badge_height', label: 'Badge height (px)', tab: 'badges', type: 'number', example: '40' }
  ];
  SOCIAL.forEach(function (n) { FIELDS.push({ key: n[0], label: n[1], tab: 'social', type: 'url', example: n[2] }); });

  var TOGGLES = [
    { key: 'social', label: 'Show social icons', tab: 'social' },
    { key: 'disclaimer', label: 'Show disclaimer', tab: 'disclaimer' },
    { key: 'banner', label: 'Show banner', tab: 'banner' },
    { key: 'extras', label: 'Show extras', tab: 'extras' },
    { key: 'badges', label: 'Show badges', tab: 'badges' }
  ];

  /* ---------- helpers ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function telHref(p) { return 'tel:' + String(p).replace(/[^\d+]/g, ''); }
  function lower(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
  function flag(v) {
    if (v === true || v === false) return v;
    var s = lower(v);
    if (s === '') return true;
    return ['yes', 'true', '1', 'y'].indexOf(s) >= 0;
  }
  function hex(v, d) { return /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(String(v || '').trim()) ? String(v).trim() : d; }

  /* Image URLs may carry their natural size as "#wh=900x400" (added by the portal).
     That lets every image get exact width + height attributes. */
  function imgInfo(url) {
    var u = String(url || ''), m = /#wh=(\d+)x(\d+)$/.exec(u);
    return { src: u.replace(/#wh=\d+x\d+$/, ''), w: m ? +m[1] : 0, h: m ? +m[2] : 0 };
  }
  function img(url, w, h, alt, extra) {
    var i = imgInfo(url);
    if (w && !h) h = i.w ? Math.round(w * i.h / i.w) : 0;
    if (h && !w) w = i.h ? Math.round(h * i.w / i.h) : 0;
    return '<img src="' + esc(i.src) + '" alt="' + esc(alt || '') + '"' + (w ? ' width="' + w + '"' : '') + (h ? ' height="' + h + '"' : '') +
      ' border="0" style="display:' + (extra && extra.inline ? 'inline-block' : 'block') + ';' + (w ? 'width:' + w + 'px;' : '') + (h ? 'height:' + h + 'px;' : 'height:auto;') +
      'max-width:' + (w || 600) + 'px;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;vertical-align:middle;">';
  }
  function link(href, inner) { return href ? '<a href="' + esc(href) + '" target="_blank" style="text-decoration:none;border:0;">' + inner + '</a>' : inner; }

  var TB = '<table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;';
  function tbl(rows, width, extraStyle) {
    return TB + (width ? 'width:' + width + 'px;' : '') + (extraStyle || '') + '"' + (width ? ' width="' + width + '"' : '') + '>' + rows + '</table>';
  }
  function txt(size, lh, color, extra) {
    return FONT + 'font-size:' + size + 'px;line-height:' + lh + 'px;mso-line-height-rule:exactly;color:' + color + ';' + (extra || '');
  }
  function tr(content, style, attrs) { return '<tr><td ' + (attrs || '') + ' style="' + style + '">' + content + '</td></tr>'; }
  function gapRow(h) { return '<tr><td height="' + h + '" style="height:' + h + 'px;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>'; }
  function vline(color, w) { return '<td width="' + w + '" bgcolor="' + color + '" style="width:' + w + 'px;background-color:' + color + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td>'; }
  function hline(color, w, h) { return tbl('<tr>' + '<td width="' + w + '" height="' + h + '" bgcolor="' + color + '" style="width:' + w + 'px;height:' + h + 'px;background-color:' + color + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>'); }
  function spacerCell(w) { return '<td width="' + w + '" style="width:' + w + 'px;font-size:1px;line-height:1px;">&nbsp;</td>'; }

  /** General settings -> branch overrides -> person overrides */
  function resolveSettings(global, overrides, branch, email) {
    var out = {}, k;
    for (k in global) out[k] = global[k];
    if (!out.office_locations && out.branches) out.office_locations = out.branches;
    (overrides || []).forEach(function (o) {
      if (lower(o.scope) === 'branch' && lower(o.target) === lower(branch) && o.value !== '') out[o.field] = o.value;
    });
    (overrides || []).forEach(function (o) {
      if (lower(o.scope) === 'person' && lower(o.target) === lower(email) && o.value !== '') out[o.field] = o.value;
    });
    return out;
  }

  /* ---------- icons ---------- */
  function iconDesign(s) {
    return {
      shape: ['circle', 'rounded', 'square', 'plain'].indexOf(lower(s.icon_shape)) >= 0 ? lower(s.icon_shape) : 'circle',
      mode: lower(s.icon_color_mode) === 'brand' ? 'brand' : 'single',
      bg: hex(s.icon_bg, hex(s.brand_color, '#477ccc')).toLowerCase(),
      fg: hex(s.icon_fg, '#ffffff').toLowerCase()
    };
  }
  /** Folder name for a custom icon set, e.g. "circle-477ccc-ffffff" or "rounded-brand-ffffff" */
  function iconSetKey(s) {
    if (lower(s.icon_style) === 'official') return 'official-color-ffffff';
    var d = iconDesign(s);
    return [d.shape, d.mode === 'brand' ? 'brand' : d.bg.replace('#', ''), d.fg.replace('#', '')].join('-');
  }
  function iconBase(s) {
    var own = String(s.icon_base_url || '').trim();
    if (own && !/htmlsig-assets\/(round|square|rounded|polygon|grey)\/?$/.test(own)) return own.replace(/\/?$/, '/');
    var st = lower(s.icon_style) || 'round';
    if (st === 'custom' || st === 'official') return String(s.asset_base_url || '').replace(/\/?$/, '/') + 'icons/' + iconSetKey(s) + '/';
    return ASSETS + (['round', 'square', 'rounded', 'polygon', 'grey'].indexOf(st) >= 0 ? st : 'round') + '/';
  }
  function socialIcons(s, size) {
    size = parseInt(s.icon_size, 10) || size;
    var gap = parseInt(s.icon_gap, 10); if (isNaN(gap)) gap = 6;
    var base = iconBase(s), list = SOCIAL.filter(function (n) { return s[n[0]]; });
    if (!list.length) return '';
    var cells = list.map(function (n, i) {
      var src = (ICONDATA && ICONDATA[n[0]]) || base + n[0] + '.png';
      var im = '<img src="' + esc(src) + '" alt="' + esc(n[1]) + '" width="' + size + '" height="' + size + '" border="0" style="display:block;width:' + size + 'px;height:' + size + 'px;border:0;">';
      return '<td width="' + size + '" style="width:' + size + 'px;padding:0;">' + link(s[n[0]], im) + '</td>' + (i < list.length - 1 && gap ? spacerCell(gap) : '');
    }).join('');
    return tbl('<tr>' + cells + '</tr>');
  }

  /* ---------- contact details ---------- */
  function contactRows(p, s) {
    var office = p.office || s.office_phone, web = s.website || '';
    var webLabel = web.replace(/^https?:\/\//, '').replace(/\/$/, '');
    var rows = [];
    if (p.mobile) rows.push(['Mobile', 'M', '<a href="' + telHref(p.mobile) + '" style="color:' + INK + ';text-decoration:none;white-space:nowrap;">' + esc(p.mobile) + '</a>']);
    if (office) rows.push(['Office', 'O', '<a href="' + telHref(office) + '" style="color:' + INK + ';text-decoration:none;white-space:nowrap;">' + esc(office) + '</a>']);
    if (s.fax) rows.push(['Fax', 'F', '<span style="color:' + INK + ';white-space:nowrap;">' + esc(s.fax) + '</span>']);
    if (p.email) rows.push(['Email', 'E', '<a href="mailto:' + esc(p.email) + '" style="color:' + LINK + ';text-decoration:none;white-space:nowrap;">' + esc(p.email) + '</a>']);
    if (web) rows.push(['Web', 'W', '<a href="' + esc(web) + '" target="_blank" style="color:' + LINK + ';text-decoration:none;white-space:nowrap;">' + esc(webLabel) + '</a>']);
    if (s.address || s.address2) rows.push(['Address', 'A', '<span style="color:' + INK + ';">' + [s.address, s.address2].filter(Boolean).map(esc).join('<br>') + '</span>']);
    return rows;
  }
  /** Two-column table: labels line up exactly in Gmail, Outlook and Apple Mail */
  function contactTable(p, s, dflt, labelColor, bold, subset) {
    var m = LABELMODE || dflt, rows = subset || contactRows(p, s);
    return tbl(rows.map(function (r) {
      var lab = m === 'none' ? '' : '<td valign="top" style="' + txt(12, 18, labelColor || MUTE, 'padding:0 8px 0 0;white-space:nowrap;' + (bold ? 'font-weight:bold;' : '')) + '">' + (m === 'letters' ? r[1] : r[0] + ':') + '</td>';
      return '<tr>' + lab + '<td valign="top" style="' + txt(12, 18, INK) + '">' + r[2] + '</td></tr>';
    }).join(''));
  }
  /** One-line contact list for the compact / centred / minimal layouts */
  function contactInline(p, s, dfltLabel, dfltSep) {
    var m = LABELMODE || dfltLabel, sm = SEPMODE || dfltSep;
    var sepHtml = sm === 'none' ? '&nbsp;&nbsp;&nbsp;' : '&nbsp;&nbsp;<span style="color:' + (sm === '|' ? '#c8c8c8' : MUTE) + ';">' + (sm === '|' ? '|' : '/') + '</span>&nbsp;&nbsp;';
    return contactRows(p, s).map(function (r) {
      return '<span style="white-space:nowrap;">' + (m === 'none' ? '' : '<span style="color:' + MUTE + ';">' + (m === 'letters' ? r[1] : r[0] + ':') + '</span>&nbsp;') + r[2] + '</span>';
    }).join(sepHtml);
  }
  function locations(s) { return String(s.office_locations || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean).join(' | '); }
  function logo(s, w) {
    w = parseInt(s.logo_width, 10) || w;
    return s.logo_url ? link(s.website, img(s.logo_url, w, 0, s.company_name)) : '';
  }
  function nameRow(p, size) { return tr(esc(p.name), txt(size, size + 4, INK, 'font-weight:bold;')); }
  function titleRow(p, c, padBottom) { return p.designation ? tr(esc(p.designation), txt(12, 16, c, 'padding:0 0 ' + (padBottom || 0) + 'px 0;')) : ''; }
  function socialRow(s, o, top) { var ic = o.social ? socialIcons(s, 20) : ''; return ic ? gapRow(top || 8) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : ''; }

  /* ---------- layouts ---------- */
  function tplClassic(p, s, o, c) {
    var locs = locations(s);
    var details = tbl(
      nameRow(p, 16) + titleRow(p, c, 8) +
      '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
      tr(esc(s.company_name), txt(12, 16, INK, 'font-weight:bold;padding:8px 0 2px 0;')) +
      (locs ? tr(esc(locs), txt(10, 14, MUTE)) : '') + socialRow(s, o));
    return tbl('<tr>' + (s.logo_url ? '<td valign="middle" style="padding:0 16px 0 0;">' + logo(s, 140) + '</td>' + vline(c, 3) + spacerCell(16) : '') +
      '<td valign="top" style="padding:0;">' + details + '</td></tr>');
  }

  function tplCompact(p, s, o, c) {
    var head = '<b style="font-size:14px;">' + esc(p.name) + '</b>' + (p.designation ? '&nbsp;&nbsp;<span style="color:' + c + ';">' + esc(p.designation) + '</span>' : '');
    return tbl('<tr>' + vline(c, 3) + spacerCell(12) + '<td valign="top" style="padding:0;">' + tbl(
      tr(head, txt(13, 18, INK)) + tr(esc(s.company_name), txt(12, 17, MUTE, 'padding:0 0 4px 0;')) +
      tr(contactInline(p, s, 'letters', '|'), txt(12, 17, INK)) + socialRow(s, o, 6)) + '</td></tr>');
  }

  function tplPhoto(p, s, o, c) {
    var initials = String(p.name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w[0].toUpperCase(); }).join('');
    var pic = p.photo ? img(p.photo, 88, 88, p.name).replace('border:0;', 'border:0;border-radius:44px;')
      : tbl('<tr><td width="88" height="88" align="center" valign="middle" bgcolor="' + c + '" style="width:88px;height:88px;border-radius:44px;background-color:' + c + ';' + txt(30, 88, '#ffffff', 'font-weight:bold;') + '">' + esc(initials) + '</td></tr>');
    var lg = logo(s, 90), ic = o.social ? socialIcons(s, 20) : '';
    var foot = (lg || ic) ? tbl('<tr>' + (lg ? '<td valign="middle" style="padding:0;">' + lg + '</td>' : '') + (lg && ic ? spacerCell(14) : '') + (ic ? '<td valign="middle" style="padding:0;">' + ic + '</td>' : '') + '</tr>') : '';
    return tbl('<tr><td valign="top" width="88" style="width:88px;padding:0;">' + pic + '</td>' + spacerCell(18) + '<td valign="top" style="padding:0;">' + tbl(
      nameRow(p, 17) + titleRow(p, c) + tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:0 0 8px 0;')) +
      '<tr><td style="padding:0 0 8px 0;">' + hline(c, 40, 2) + '</td></tr>' +
      '<tr><td style="padding:0 0 10px 0;">' + contactTable(p, s, 'text') + '</td></tr>' +
      (foot ? '<tr><td style="padding:0;">' + foot + '</td></tr>' : '')) + '</td></tr>');
  }

  function tplHeader(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 20) : '';
    return tbl(
      '<tr><td bgcolor="' + c + '" style="background-color:' + c + ';padding:12px 16px;">' + tbl(
        tr(esc(p.name), txt(17, 21, '#ffffff', 'font-weight:bold;')) +
        (p.designation ? tr(esc(p.designation), txt(12, 16, '#ffffff')) : '')) + '</td></tr>' +
      '<tr><td style="padding:12px 16px;border-left:1px solid #e3e6eb;border-right:1px solid #e3e6eb;border-bottom:1px solid #e3e6eb;">' +
        tbl('<tr><td valign="top" style="padding:0;">' + tbl('<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' + (ic ? gapRow(8) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : '')) + '</td>' +
          (s.logo_url ? '<td valign="middle" align="right" width="120" style="width:120px;padding:0 0 0 12px;">' + logo(s, 110) + '</td>' : '') + '</tr>', 446) +
      '</td></tr>', 480);
  }

  function tplMinimal(p, s, o) {
    return tbl(tr('--', txt(12, 16, MUTE, 'padding:0 0 6px 0;')) + tr('<b>' + esc(p.name) + '</b>', txt(14, 18, INK)) +
      tr(esc([p.designation, s.company_name].filter(Boolean).join(', ')), txt(12, 17, MUTE, 'padding:0 0 4px 0;')) +
      tr(contactInline(p, s, 'none', '/'), txt(12, 17, INK)) + socialRow(s, o, 6));
  }

  function tplCentered(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 20) : '';
    return tbl(
      (s.logo_url ? '<tr><td align="center" style="padding:0 0 8px 0;">' + tbl('<tr><td>' + logo(s, 120) + '</td></tr>') + '</td></tr>' : '') +
      tr(esc(p.name), txt(16, 20, INK, 'font-weight:bold;'), 'align="center"') +
      (p.designation ? tr(esc(p.designation), txt(12, 16, c, 'padding:0 0 8px 0;'), 'align="center"') : '') +
      '<tr><td align="center" style="padding:0 0 8px 0;">' + hline(c, 60, 2) + '</td></tr>' +
      tr(contactInline(p, s, 'none', '|'), txt(12, 18, INK), 'align="center"') +
      (ic ? '<tr><td align="center" style="padding:8px 0 0 0;">' + ic + '</td></tr>' : ''), 480);
  }

  function tplCard(p, s, o, c) {
    var locs = locations(s), ic = o.social ? socialIcons(s, 20) : '';
    return tbl(
      '<tr><td style="padding:16px 16px 12px 16px;border-top:1px solid #e3e6eb;border-left:1px solid #e3e6eb;border-right:1px solid #e3e6eb;">' +
        tbl('<tr><td valign="top" style="padding:0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 10) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>') + '</td>' +
          (s.logo_url ? '<td valign="top" align="right" width="120" style="width:120px;padding:0 0 0 12px;">' + logo(s, 110) + '</td>' : '') + '</tr>', 446) +
      '</td></tr>' +
      '<tr><td style="padding:0 16px 12px 16px;border-left:1px solid #e3e6eb;border-right:1px solid #e3e6eb;">' + tbl(
        tr(esc(s.company_name), txt(10, 14, INK, 'font-weight:bold;')) + (locs ? tr(esc(locs), txt(10, 14, MUTE)) : '') +
        (ic ? gapRow(8) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : '')) + '</td></tr>' +
      '<tr><td height="4" bgcolor="' + c + '" style="height:4px;background-color:' + c + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>', 480);
  }

  function tplLogoTop(p, s, o, c) {
    return tbl(
      (s.logo_url ? '<tr><td colspan="3" style="padding:0 0 10px 0;">' + logo(s, 150) + '</td></tr>' : '') +
      '<tr><td colspan="3" height="2" bgcolor="' + c + '" style="height:2px;background-color:' + c + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>' +
      '<tr><td valign="top" width="220" style="width:220px;padding:10px 0 0 0;">' + tbl(nameRow(p, 16) + titleRow(p, c) +
        tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:4px 0 0 0;')) + socialRow(s, o)) + '</td>' + spacerCell(16) +
      '<td valign="top" style="padding:10px 0 0 0;">' + contactTable(p, s, 'text') + '</td></tr>', 480);
  }

  function tplSplit(p, s, o, c) {
    var lg = logo(s, 100), ic = o.social ? socialIcons(s, 20) : '';
    return tbl('<tr><td valign="middle" style="padding:0;">' + tbl(nameRow(p, 18) + titleRow(p, c, lg ? 8 : 0) + (lg ? '<tr><td style="padding:0;">' + lg + '</td></tr>' : '')) + '</td>' +
      spacerCell(16) + vline(c, 2) + spacerCell(16) +
      '<td valign="middle" style="padding:0;">' + tbl('<tr><td style="padding:0;">' + contactTable(p, s, 'letters', c, true) + '</td></tr>' +
        (ic ? gapRow(8) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : '')) + '</td></tr>');
  }

  /* ---------- more layouts ---------- */
  function initialsOf(p) { return String(p.name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w[0].toUpperCase(); }).join(''); }
  function avatar(p, c, size, round) {
    var r = round ? Math.round(size / 2) : 8;
    return p.photo ? img(p.photo, size, size, p.name).replace('border:0;', 'border:0;border-radius:' + r + 'px;')
      : tbl('<tr><td width="' + size + '" height="' + size + '" align="center" valign="middle" bgcolor="' + c + '" style="width:' + size + 'px;height:' + size + 'px;border-radius:' + r + 'px;background-color:' + c + ';' + txt(Math.round(size * 0.34), size, '#ffffff', 'font-weight:bold;') + '">' + esc(initialsOf(p)) + '</td></tr>');
  }
  function twoCols(p, s) {
    var rows = contactRows(p, s), half = Math.ceil(rows.length / 2);
    return tbl('<tr><td valign="top" style="padding:0;">' + contactTable(p, s, 'text', null, false, rows.slice(0, half)) + '</td>' + spacerCell(24) +
      '<td valign="top" style="padding:0;">' + contactTable(p, s, 'text', null, false, rows.slice(half)) + '</td></tr>');
  }
  function fullRule(color, w, colspan) { return '<tr><td' + (colspan ? ' colspan="' + colspan + '"' : '') + ' height="1" bgcolor="' + color + '" style="height:1px;background-color:' + color + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>'; }

  // Corporate: name + logo, rule, contacts in two columns, company bar
  function tplExecutive(p, s, o, c) {
    var locs = locations(s), ic = o.social ? socialIcons(s, 20) : '';
    return tbl(
      '<tr><td valign="bottom" style="padding:0 0 8px 0;">' + tbl(nameRow(p, 18) + titleRow(p, c)) + '</td>' +
      '<td valign="bottom" align="right" style="padding:0 0 8px 0;">' + logo(s, 110) + '</td></tr>' +
      fullRule('#dfe3ea', 480, 2) +
      '<tr><td colspan="2" style="padding:8px 0;">' + twoCols(p, s) + '</td></tr>' +
      '<tr><td colspan="2" bgcolor="' + c + '" style="background-color:' + c + ';padding:6px 10px;' + txt(11, 15, '#ffffff') + '"><b>' + esc(s.company_name) + '</b>' + (locs ? '&nbsp;&nbsp;' + esc(locs) : '') + '</td></tr>' +
      (ic ? '<tr><td colspan="2" style="padding:8px 0 0 0;">' + ic + '</td></tr>' : ''), 480);
  }
  // Corporate: details with logo, coloured strip with company and website
  function tplBrandBar(p, s, o, c) {
    var web = String(s.website || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
    return tbl(
      '<tr><td valign="top" style="padding:0 0 10px 0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 8) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' + socialRow(s, o)) + '</td>' +
      '<td valign="top" align="right" width="130" style="width:130px;padding:0 0 10px 12px;">' + logo(s, 120) + '</td></tr>' +
      '<tr><td bgcolor="' + c + '" style="background-color:' + c + ';padding:8px 12px;' + txt(12, 16, '#ffffff', 'font-weight:bold;') + '">' + esc(s.company_name) + '</td>' +
      '<td bgcolor="' + c + '" align="right" style="background-color:' + c + ';padding:8px 12px;' + txt(12, 16, '#ffffff') + '">' + (web ? '<a href="' + esc(s.website) + '" target="_blank" style="color:#ffffff;text-decoration:none;">' + esc(web) + '</a>' : '') + '</td></tr>', 480);
  }
  // Corporate: light panel with accent edge
  function tplBoxed(p, s, o, c) {
    return tbl('<tr>' + vline(c, 4) + '<td bgcolor="#f5f7fa" style="background-color:#f5f7fa;padding:14px 16px;">' +
      tbl('<tr><td valign="top" style="padding:0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 8) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
        tr(esc(s.company_name), txt(12, 16, INK, 'font-weight:bold;padding:8px 0 0 0;')) + socialRow(s, o)) + '</td>' +
        (s.logo_url ? '<td valign="top" align="right" style="padding:0 0 0 16px;">' + logo(s, 110) + '</td>' : '') + '</tr>', 440) +
      '</td></tr>', 480);
  }
  // Professional: stacked, accent rule, logo and icons at the bottom
  function tplStacked(p, s, o, c) {
    var lg = logo(s, 90), ic = o.social ? socialIcons(s, 20) : '';
    return tbl(nameRow(p, 17) + tr(esc([p.designation, s.company_name].filter(Boolean).join('  |  ')), txt(12, 16, c)) +
      '<tr><td style="padding:8px 0;">' + hline(c, 40, 2) + '</td></tr>' +
      '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
      ((lg || ic) ? '<tr><td style="padding:10px 0 0 0;">' + tbl('<tr>' + (lg ? '<td valign="middle" style="padding:0;">' + lg + '</td>' : '') + (lg && ic ? spacerCell(14) : '') + (ic ? '<td valign="middle" style="padding:0;">' + ic + '</td>' : '') + '</tr>') + '</td></tr>' : ''));
  }
  // Professional: details left, logo right
  function tplRightLogo(p, s, o, c) {
    var locs = locations(s);
    return tbl('<tr><td valign="top" style="padding:0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 8) +
      '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
      tr(esc(s.company_name), txt(12, 16, INK, 'font-weight:bold;padding:8px 0 2px 0;')) + (locs ? tr(esc(locs), txt(10, 14, MUTE)) : '') + socialRow(s, o)) + '</td>' +
      (s.logo_url ? spacerCell(16) + vline(c, 3) + spacerCell(16) + '<td valign="middle" style="padding:0;">' + logo(s, 130) + '</td>' : '') + '</tr>');
  }
  // Professional: header line with icons, contacts in two columns
  function tplTwoColumn(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 18) : '';
    return tbl('<tr><td valign="bottom" style="padding:0 0 6px 0;">' + tbl(nameRow(p, 16) + tr(esc([p.designation, s.company_name].filter(Boolean).join(', ')), txt(12, 16, c))) + '</td>' +
      '<td valign="bottom" align="right" style="padding:0 0 6px 0;">' + ic + '</td></tr>' +
      '<tr><td colspan="2" style="padding:0;">' + hline(c, 480, 2) + '</td></tr>' +
      '<tr><td colspan="2" style="padding:8px 0 0 0;">' + twoCols(p, s) + '</td></tr>' +
      (s.logo_url ? '<tr><td colspan="2" style="padding:10px 0 0 0;">' + logo(s, 100) + '</td></tr>' : ''), 480);
  }
  // Personal: round photo on top, everything centred
  function tplPhotoCenter(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 22) : '';
    return tbl('<tr><td align="center" style="padding:0 0 8px 0;">' + avatar(p, c, 80, true) + '</td></tr>' +
      tr(esc(p.name), txt(17, 21, INK, 'font-weight:bold;'), 'align="center"') +
      (p.designation ? tr(esc(p.designation), txt(12, 16, c), 'align="center"') : '') +
      tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:0 0 8px 0;'), 'align="center"') +
      tr(contactInline(p, s, 'none', '|'), txt(12, 18, INK), 'align="center"') +
      (ic ? '<tr><td align="center" style="padding:10px 0 0 0;">' + ic + '</td></tr>' : ''), 480);
  }
  // Personal: large photo, details beside it
  function tplPhotoLarge(p, s, o, c) {
    return tbl('<tr><td valign="top" width="110" style="width:110px;padding:0;">' + avatar(p, c, 110, false) + '</td>' + spacerCell(18) +
      '<td valign="top" style="padding:0;">' + tbl(nameRow(p, 18) + titleRow(p, c) + tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:0 0 8px 0;')) +
        '<tr><td style="padding:0;">' + contactTable(p, s, 'letters', c, true) + '</td></tr>' + socialRow(s, o)) + '</td></tr>');
  }
  // Personal: warm and simple, name in accent colour, big icons
  function tplPersonal(p, s, o, c) {
    return tbl(tr(esc(p.name), txt(20, 24, c, 'font-weight:bold;')) +
      (p.designation ? tr(esc(p.designation), txt(13, 18, INK)) : '') +
      tr(contactInline(p, s, 'letters', '/'), txt(12, 18, INK, 'padding:6px 0 0 0;')) +
      (o.social && socialIcons(s, 26) ? gapRow(10) + '<tr><td style="padding:0;">' + socialIcons(s, 26) + '</td></tr>' : ''));
  }
  // Creative: initials monogram block
  function tplMonogram(p, s, o, c) {
    var mono = tbl('<tr><td width="88" height="88" align="center" valign="middle" bgcolor="' + c + '" style="width:88px;height:88px;background-color:' + c + ';' + txt(32, 88, '#ffffff', 'font-weight:bold;letter-spacing:1px;') + '">' + esc(initialsOf(p)) + '</td></tr>');
    return tbl('<tr><td valign="top" style="padding:0;">' + mono + (s.logo_url ? '<table border="0" cellpadding="0" cellspacing="0" role="presentation"><tr><td style="padding:8px 0 0 0;">' + logo(s, 88) + '</td></tr></table>' : '') + '</td>' + spacerCell(16) +
      '<td valign="top" style="padding:0;">' + tbl(nameRow(p, 17) + titleRow(p, c, 6) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
        tr(esc(s.company_name), txt(11, 15, MUTE, 'padding:6px 0 0 0;')) + socialRow(s, o)) + '</td></tr>');
  }
  // Creative: very large name in accent colour
  function tplBoldName(p, s, o, c) {
    return tbl(tr(esc(p.name), txt(26, 30, c, 'font-weight:bold;')) +
      tr(esc([p.designation, s.company_name].filter(Boolean).join(', ')), txt(13, 18, INK)) +
      '<tr><td style="padding:8px 0;">' + hline(c, 60, 4) + '</td></tr>' +
      tr(contactInline(p, s, 'letters', '/'), txt(12, 18, INK)) + socialRow(s, o) +
      (s.logo_url ? '<tr><td style="padding:10px 0 0 0;">' + logo(s, 100) + '</td></tr>' : ''));
  }
  // Creative: coloured panel with name, white panel with details
  function tplColorSplit(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 20) : '';
    return tbl('<tr><td valign="top" width="170" bgcolor="' + c + '" style="width:170px;background-color:' + c + ';padding:14px;">' +
        tbl(tr(esc(p.name), txt(16, 20, '#ffffff', 'font-weight:bold;')) + (p.designation ? tr(esc(p.designation), txt(12, 16, '#ffffff', 'padding:2px 0 0 0;')) : '') +
          (ic ? gapRow(12) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : '')) + '</td>' +
      '<td valign="top" style="padding:14px;border-top:1px solid #e3e6eb;border-right:1px solid #e3e6eb;border-bottom:1px solid #e3e6eb;">' +
        tbl('<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' + tr(esc(s.company_name), txt(12, 16, INK, 'font-weight:bold;padding:8px 0 0 0;')) +
          (s.logo_url ? '<tr><td style="padding:8px 0 0 0;">' + logo(s, 100) + '</td></tr>' : '')) + '</td></tr>', 480);
  }
  // Minimal: two lines of text
  function tplOneLine(p, s, o, c) {
    return tbl(tr('<b>' + esc(p.name) + '</b>' + (p.designation ? '&nbsp;&nbsp;|&nbsp;&nbsp;' + esc(p.designation) : '') + (s.company_name ? '&nbsp;&nbsp;|&nbsp;&nbsp;<span style="color:' + c + ';">' + esc(s.company_name) + '</span>' : ''), txt(13, 18, INK)) +
      tr(contactInline(p, s, 'letters', '|'), txt(12, 18, INK)) + socialRow(s, o, 6));
  }
  // Minimal: one colour, no accents
  function tplMono(p, s, o) {
    return tbl(tr(esc(p.name), txt(13, 18, INK, 'font-weight:bold;')) + (p.designation ? tr(esc(p.designation), txt(12, 17, INK)) : '') +
      tr(esc(s.company_name), txt(12, 17, INK, 'padding:0 0 6px 0;')) +
      '<tr><td style="padding:0;">' + contactTable(p, s, 'letters', INK) + '</td></tr>' + socialRow(s, o, 6));
  }

  var LAYOUTS = { classic: tplClassic, compact: tplCompact, photo: tplPhoto, header: tplHeader, minimal: tplMinimal,
                  centered: tplCentered, card: tplCard, logotop: tplLogoTop, split: tplSplit,
                  executive: tplExecutive, brandbar: tplBrandBar, boxed: tplBoxed, stacked: tplStacked, rightlogo: tplRightLogo,
                  twocol: tplTwoColumn, photocenter: tplPhotoCenter, photolarge: tplPhotoLarge, personal: tplPersonal,
                  monogram: tplMonogram, boldname: tplBoldName, colorsplit: tplColorSplit, oneline: tplOneLine, mono: tplMono };
  var CATEGORIES = [['corporate', 'Corporate'], ['professional', 'Professional'], ['personal', 'Personal'], ['creative', 'Creative'], ['minimal', 'Minimal']];
  var TEMPLATES = [
    { id: 'classic', label: 'Classic', cat: 'corporate' }, { id: 'executive', label: 'Executive', cat: 'corporate' },
    { id: 'brandbar', label: 'Brand bar', cat: 'corporate' }, { id: 'header', label: 'Header band', cat: 'corporate' },
    { id: 'boxed', label: 'Boxed', cat: 'corporate' }, { id: 'logotop', label: 'Logo on top', cat: 'corporate' },
    { id: 'card', label: 'Card', cat: 'professional' }, { id: 'split', label: 'Split', cat: 'professional' },
    { id: 'stacked', label: 'Stacked', cat: 'professional' }, { id: 'rightlogo', label: 'Logo right', cat: 'professional' },
    { id: 'twocol', label: 'Two column', cat: 'professional' },
    { id: 'photo', label: 'Photo', cat: 'personal' }, { id: 'photocenter', label: 'Photo centred', cat: 'personal' },
    { id: 'photolarge', label: 'Large photo', cat: 'personal' }, { id: 'personal', label: 'Personal', cat: 'personal' },
    { id: 'monogram', label: 'Monogram', cat: 'creative' }, { id: 'boldname', label: 'Bold name', cat: 'creative' },
    { id: 'colorsplit', label: 'Colour split', cat: 'creative' }, { id: 'centered', label: 'Centred', cat: 'creative' },
    { id: 'compact', label: 'Compact', cat: 'minimal' }, { id: 'minimal', label: 'Minimal', cat: 'minimal' },
    { id: 'oneline', label: 'One line', cat: 'minimal' }, { id: 'mono', label: 'Mono', cat: 'minimal' }
  ];

  function options(p) {
    return {
      template: LAYOUTS.hasOwnProperty(p.template) ? p.template : 'classic',
      banner: flag(p.banner), social: flag(p.social), disclaimer: flag(p.disclaimer), extras: flag(p.extras), badges: flag(p.badges)
    };
  }

  /* ---------- blocks under the main layout ---------- */
  function extrasBlock(s, c) {
    var rows = '';
    if (s.tagline) rows += tr('<i>' + esc(s.tagline) + '</i>', txt(12, 17, MUTE, 'padding:0 0 8px 0;'));
    var btns = [];
    function imgBtn(href, file, w, alt) { return link(href, '<img src="' + ASSETS + file + '" alt="' + alt + '" width="' + w + '" height="37" border="0" style="display:block;width:' + w + 'px;height:37px;border:0;">'); }
    if (s.app_apple) btns.push(imgBtn(s.app_apple, 'app-icon/apple.png', 119, 'Download on the App Store'));
    if (s.app_google) btns.push(imgBtn(s.app_google, 'app-icon/google.png', 119, 'Get it on Google Play'));
    if (s.app_amazon) btns.push(imgBtn(s.app_amazon, 'app-icon/amazon.png', 119, 'Available at Amazon Appstore'));
    if (s.calendar_link) btns.push(imgBtn(s.calendar_link, 'calendar/schedule-a-meeting.png', 100, 'Schedule a meeting'));
    var wa = String(s.whatsapp_number || '').replace(/[^\d]/g, '');
    if (wa) btns.push(imgBtn('https://wa.me/' + wa + (s.whatsapp_message ? '?text=' + encodeURIComponent(s.whatsapp_message) : ''), 'whatsapp/whatsapp-contact_me.png', 119, 'Contact me on WhatsApp'));
    if (btns.length) rows += '<tr><td style="padding:0 0 8px 0;">' + tbl('<tr>' + btns.map(function (b, i) { return '<td style="padding:0;">' + b + '</td>' + (i < btns.length - 1 ? spacerCell(8) : ''); }).join('') + '</tr>') + '</td></tr>';
    if (s.cta_text && s.cta_link) {
      rows += '<tr><td style="padding:0 0 8px 0;">' + TB + 'border-collapse:separate;"><tr><td bgcolor="' + c + '" style="background-color:' + c + ';border-radius:4px;padding:7px 16px;">' +
        '<a href="' + esc(s.cta_link) + '" target="_blank" style="' + txt(12, 16, '#ffffff', 'font-weight:bold;text-decoration:none;display:inline-block;') + '">' + esc(s.cta_text) + '</a></td></tr></table></td></tr>';
    }
    if (s.eco_note) rows += tr(esc(s.eco_note), txt(10, 14, '#3a7d44'));
    return rows ? '<tr><td style="padding:12px 0 0 0;">' + tbl(rows) + '</td></tr>' : '';
  }

  function badgesBlock(s) {
    var h = parseInt(s.badge_height, 10) || 40, cells = [];
    for (var i = 1; i <= 4; i++) {
      var im = s['badge' + i + '_img'];
      if (im) cells.push('<td valign="middle" style="padding:0;">' + link(s['badge' + i + '_link'], img(im, 0, h, 'Badge')) + '</td>');
    }
    return cells.length ? '<tr><td style="padding:12px 0 0 0;">' + tbl('<tr>' + cells.join(spacerCell(10)) + '</tr>') + '</td></tr>' : '';
  }

  /** p: person, s: resolved settings, extra: { iconData } (portal preview only). Branch is never printed. */
  function render(p, s, extra) {
    var o = options(p);
    var c = hex(s.brand_color, '#477ccc');
    INK = hex(s.text_color, '#212121');
    LINK = hex(s.link_color, c);
    MUTE = hex(s.secondary_color, '#6b6b6b');
    FONT = 'font-family:' + (String(s.font_family || '').replace(/["<>;{}]/g, '').trim() || DEFAULT_FONT) + ';';
    LABELMODE = ['text', 'letters', 'none'].indexOf(lower(s.phone_labels)) >= 0 ? lower(s.phone_labels) : '';
    SEPMODE = ['/', '|', 'none'].indexOf(lower(s.separator)) >= 0 ? lower(s.separator) : '';
    ICONDATA = (extra && extra.iconData) || null;
    var k = { small: 0.9, large: 1.12 }[lower(s.font_size)] || 1;

    var body = LAYOUTS[o.template](p, s, o, c);
    var bw = parseInt(s.banner_width, 10) || 450;
    var banner = o.banner && s.banner_url ? '<tr><td style="padding:14px 0 0 0;">' + link(s.banner_link, img(s.banner_url, bw, 0, s.company_name)) + '</td></tr>' : '';
    var w = Math.max(480, bw);
    var disc = o.disclaimer && s.disclaimer
      ? '<tr><td style="padding:14px 0 0 0;">' + tbl('<tr><td style="border-top:1px solid #dddddd;padding:10px 0 0 0;' + txt(9, 12, '#8a8a8a') + '">' + esc(s.disclaimer) + '</td></tr>', w) + '</td></tr>' : '';

    var html = tbl('<tr><td align="left" style="padding:0;">' + body + '</td></tr>' +
      (o.extras ? extrasBlock(s, c) : '') + banner + (o.badges ? badgesBlock(s) : '') + disc, w, FONT + '-webkit-text-size-adjust:none;-ms-text-size-adjust:none;');
    ICONDATA = null;
    if (k !== 1) html = html.replace(/(font-size|line-height):(\d+)px/g, function (m, prop, n) {
      n = +n; return n <= 1 ? m : prop + ':' + Math.round(n * k) + 'px';
    });
    return html;
  }

  /** The downloadable file: only the signature, so Ctrl+A copies exactly that */
  function fileHtml(p, s) {
    return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width">' +
      '<title>Email signature - ' + esc(p.name) + '</title></head><body style="margin:0;padding:16px;background:#ffffff;">' + render(p, s) + '</body></html>';
  }

  return { FIELDS: FIELDS, TABS: TABS, TOGGLES: TOGGLES, TEMPLATES: TEMPLATES, CATEGORIES: CATEGORIES, SOCIAL: SOCIAL,
           resolveSettings: resolveSettings, render: render, fileHtml: fileHtml, esc: esc, flag: flag, hex: hex,
           iconDesign: iconDesign, iconSetKey: iconSetKey, imgInfo: imgInfo };
})();
