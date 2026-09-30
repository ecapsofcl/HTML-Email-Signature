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

  // Set at the start of every render() from Style settings
  var FONT = 'font-family:' + DEFAULT_FONT + ';';
  var INK = '#212121', MUTE = '#6b6b6b', LINK = '#477ccc';
  var LABELMODE = '', SEPMODE = '', ICONDATA = null;
  var MAX_W = 600, SW = 600;   // 600px = widest size every email program shows without scrolling

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
    { key: 'banner_width',     label: 'Display width (px)',      tab: 'banner', type: 'number', example: 'Empty = full signature width', hint: 'Leave empty to fill the signature width (600 px). Never wider than the signature.' },

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
    { key: 'sig_width',        label: 'Signature width',         tab: 'style', type: 'select', options: [['600', '600 px (maximum, recommended)'], ['560', '560 px'], ['520', '520 px'], ['480', '480 px']] },
    { key: 'logo_width',       label: 'Logo width (px)',         tab: 'style', type: 'number' },

    { key: 'icon_style',       label: 'Icon design',             tab: 'icons', type: 'select', options: [['official', 'Official logos (full colour)'], ['custom', 'Custom (your shape and colours)']] },
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
  FIELDS.push({ key: 'social_show', label: 'Icons to show', tab: 'social', type: 'socialpick', wide: true,
                hint: 'Tick the icons to include in the signature. Networks without a link are listed once you add their link below.' });
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
      (w ? 'min-width:' + w + 'px;max-width:' + w + 'px;' : '') + 'border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;vertical-align:middle;">';
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
    if (lower(s.icon_style) !== 'custom') return 'official-color-ffffff';   // older "standard" choices now use the official set
    var d = iconDesign(s);
    return [d.shape, d.mode === 'brand' ? 'brand' : d.bg.replace('#', ''), d.fg.replace('#', '')].join('-');
  }
  /** All images are served from your own site (asset_base_url = GitHub Pages address) */
  function assetBase(s) { return String(s.asset_base_url || '').replace(/\/?$/, '/'); }
  function iconBase(s) {
    var own = String(s.icon_base_url || '').trim();
    if (own && !/amazonaws\.com/.test(own)) return own.replace(/\/?$/, '/');
    return assetBase(s) + 'icons/' + iconSetKey(s) + '/';
  }
  function socialIcons(s, size) {
    size = parseInt(s.icon_size, 10) || size;
    var gap = parseInt(s.icon_gap, 10); if (isNaN(gap)) gap = 6;
    var show = String(s.social_show || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    var base = iconBase(s), list = SOCIAL.filter(function (n) {
      if (!s[n[0]]) return false;                                        // no link, no icon
      if (!show.length || show.indexOf('all') >= 0) return true;         // nothing chosen = every network with a link
      return show.indexOf(n[0]) >= 0;                                    // only the ticked ones ('none' matches nothing)
    });
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
          (s.logo_url ? '<td valign="middle" align="right" width="120" style="width:120px;padding:0 0 0 12px;">' + logo(s, 110) + '</td>' : '') + '</tr>', SW - 34) +
      '</td></tr>', SW);
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
      (ic ? '<tr><td align="center" style="padding:8px 0 0 0;">' + ic + '</td></tr>' : ''), SW);
  }

  function tplCard(p, s, o, c) {
    var locs = locations(s), ic = o.social ? socialIcons(s, 20) : '';
    return tbl(
      '<tr><td style="padding:16px 16px 12px 16px;border-top:1px solid #e3e6eb;border-left:1px solid #e3e6eb;border-right:1px solid #e3e6eb;">' +
        tbl('<tr><td valign="top" style="padding:0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 10) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>') + '</td>' +
          (s.logo_url ? '<td valign="top" align="right" width="120" style="width:120px;padding:0 0 0 12px;">' + logo(s, 110) + '</td>' : '') + '</tr>', SW - 34) +
      '</td></tr>' +
      '<tr><td style="padding:0 16px 12px 16px;border-left:1px solid #e3e6eb;border-right:1px solid #e3e6eb;">' + tbl(
        tr(esc(s.company_name), txt(10, 14, INK, 'font-weight:bold;')) + (locs ? tr(esc(locs), txt(10, 14, MUTE)) : '') +
        (ic ? gapRow(8) + '<tr><td style="padding:0;">' + ic + '</td></tr>' : '')) + '</td></tr>' +
      '<tr><td height="4" bgcolor="' + c + '" style="height:4px;background-color:' + c + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>', SW);
  }

  function tplLogoTop(p, s, o, c) {
    return tbl(
      (s.logo_url ? '<tr><td colspan="3" style="padding:0 0 10px 0;">' + logo(s, 150) + '</td></tr>' : '') +
      '<tr><td colspan="3" height="2" bgcolor="' + c + '" style="height:2px;background-color:' + c + ';font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>' +
      '<tr><td valign="top" width="220" style="width:220px;padding:10px 0 0 0;">' + tbl(nameRow(p, 16) + titleRow(p, c) +
        tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:4px 0 0 0;')) + socialRow(s, o)) + '</td>' + spacerCell(16) +
      '<td valign="top" style="padding:10px 0 0 0;">' + contactTable(p, s, 'text') + '</td></tr>', SW);
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
      fullRule('#dfe3ea', SW, 2) +
      '<tr><td colspan="2" style="padding:8px 0;">' + twoCols(p, s) + '</td></tr>' +
      '<tr><td colspan="2" bgcolor="' + c + '" style="background-color:' + c + ';padding:6px 10px;' + txt(11, 15, '#ffffff') + '"><b>' + esc(s.company_name) + '</b>' + (locs ? '&nbsp;&nbsp;' + esc(locs) : '') + '</td></tr>' +
      (ic ? '<tr><td colspan="2" style="padding:8px 0 0 0;">' + ic + '</td></tr>' : ''), SW);
  }
  // Corporate: details with logo, coloured strip with company and website
  function tplBrandBar(p, s, o, c) {
    var web = String(s.website || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
    return tbl(
      '<tr><td valign="top" style="padding:0 0 10px 0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 8) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' + socialRow(s, o)) + '</td>' +
      '<td valign="top" align="right" width="130" style="width:130px;padding:0 0 10px 12px;">' + logo(s, 120) + '</td></tr>' +
      '<tr><td bgcolor="' + c + '" style="background-color:' + c + ';padding:8px 12px;' + txt(12, 16, '#ffffff', 'font-weight:bold;') + '">' + esc(s.company_name) + '</td>' +
      '<td bgcolor="' + c + '" align="right" style="background-color:' + c + ';padding:8px 12px;' + txt(12, 16, '#ffffff') + '">' + (web ? '<a href="' + esc(s.website) + '" target="_blank" style="color:#ffffff;text-decoration:none;">' + esc(web) + '</a>' : '') + '</td></tr>', SW);
  }
  // Corporate: light panel with accent edge
  function tplBoxed(p, s, o, c) {
    return tbl('<tr>' + vline(c, 4) + '<td bgcolor="#f5f7fa" style="background-color:#f5f7fa;padding:14px 16px;">' +
      tbl('<tr><td valign="top" style="padding:0;">' + tbl(nameRow(p, 16) + titleRow(p, c, 8) + '<tr><td style="padding:0;">' + contactTable(p, s, 'text') + '</td></tr>' +
        tr(esc(s.company_name), txt(12, 16, INK, 'font-weight:bold;padding:8px 0 0 0;')) + socialRow(s, o)) + '</td>' +
        (s.logo_url ? '<td valign="top" align="right" style="padding:0 0 0 16px;">' + logo(s, 110) + '</td>' : '') + '</tr>', SW - 40) +
      '</td></tr>', SW);
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
      '<tr><td colspan="2" style="padding:0;">' + hline(c, SW, 2) + '</td></tr>' +
      '<tr><td colspan="2" style="padding:8px 0 0 0;">' + twoCols(p, s) + '</td></tr>' +
      (s.logo_url ? '<tr><td colspan="2" style="padding:10px 0 0 0;">' + logo(s, 100) + '</td></tr>' : ''), SW);
  }
  // Personal: round photo on top, everything centred
  function tplPhotoCenter(p, s, o, c) {
    var ic = o.social ? socialIcons(s, 22) : '';
    return tbl('<tr><td align="center" style="padding:0 0 8px 0;">' + avatar(p, c, 80, true) + '</td></tr>' +
      tr(esc(p.name), txt(17, 21, INK, 'font-weight:bold;'), 'align="center"') +
      (p.designation ? tr(esc(p.designation), txt(12, 16, c), 'align="center"') : '') +
      tr(esc(s.company_name), txt(12, 16, MUTE, 'padding:0 0 8px 0;'), 'align="center"') +
      tr(contactInline(p, s, 'none', '|'), txt(12, 18, INK), 'align="center"') +
      (ic ? '<tr><td align="center" style="padding:10px 0 0 0;">' + ic + '</td></tr>' : ''), SW);
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
          (s.logo_url ? '<tr><td style="padding:8px 0 0 0;">' + logo(s, 100) + '</td></tr>' : '')) + '</td></tr>', SW);
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

  /* =================== CUSTOM DESIGNS (Design studio) ===================
     A design is rows -> columns -> blocks. It is turned into nested tables, so it
     looks the same in Gmail, Outlook and Apple Mail. */
  var BLOCKS = {
    logo:      { label: 'Logo',            min: 40, max: 260, def: 120, unit: 'px wide' },
    photo:     { label: 'Photo',           min: 40, max: 160, def: 88,  unit: 'px' },
    name:      { label: 'Name',            min: 12, max: 34,  def: 17,  unit: 'px text', text: true, color: 'text', bold: true },
    title:     { label: 'Job title',       min: 10, max: 22,  def: 12,  unit: 'px text', text: true, color: 'accent' },
    company:   { label: 'Company name',    min: 10, max: 20,  def: 12,  unit: 'px text', text: true, color: 'text', bold: true },
    contacts:  { label: 'All contact details', min: 10, max: 16,  def: 12,  unit: 'px text' },
    mobile:    { label: 'Mobile',          min: 10, max: 18,  def: 12,  unit: 'px text', contact: 'Mobile' },
    office:    { label: 'Office phone',    min: 10, max: 18,  def: 12,  unit: 'px text', contact: 'Office' },
    email:     { label: 'Email',           min: 10, max: 18,  def: 12,  unit: 'px text', contact: 'Email' },
    web:       { label: 'Website',         min: 10, max: 18,  def: 12,  unit: 'px text', contact: 'Web' },
    fax:       { label: 'Fax',             min: 10, max: 18,  def: 12,  unit: 'px text', contact: 'Fax' },
    locations: { label: 'Office locations', min: 9, max: 14,  def: 10,  unit: 'px text', text: true, color: 'label' },
    address:   { label: 'Address',         min: 9,  max: 15,  def: 11,  unit: 'px text', text: true, color: 'label' },
    social:    { label: 'Social icons',    min: 14, max: 40,  def: 20,  unit: 'px' },
    divider:   { label: 'Line',            min: 16, max: 600, def: 60,  unit: 'px long (0 = full width)', color: 'accent' },
    spacer:    { label: 'Space',           min: 2,  max: 60,  def: 10,  unit: 'px tall' },
    text:      { label: 'Custom text',     min: 9,  max: 24,  def: 12,  unit: 'px text', text: true, color: 'text' },
    tagline:   { label: 'Tagline',         min: 10, max: 18,  def: 12,  unit: 'px text', text: true, color: 'label', italic: true },
    button:    { label: 'Button',          min: 10, max: 18,  def: 12,  unit: 'px text' },
    banner:    { label: 'Banner',          min: 120, max: 600, def: 0,  unit: 'px wide (0 = fit column)' }
  };
  function clampN(v, lo, hi, d) { v = parseFloat(v); return isNaN(v) ? d : Math.min(hi, Math.max(lo, v)); }
  function colorOf(v, c) {
    v = String(v || '');
    if (v === 'accent') return c; if (v === 'label') return MUTE; if (v === 'white') return '#ffffff'; if (v === 'text' || !v) return INK;
    return hex(v, INK);
  }
  /** Wraps block-level content so it can be centred / right-aligned reliably (table align works in every client) */
  function wrapAlign(inner, align) {
    return align === 'center' || align === 'right'
      ? '<table align="' + align + '" border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;"><tr><td style="padding:0;">' + inner + '</td></tr></table>'
      : inner;
  }
  function blockRow(p, s, o, c, b, align, colW) {
    colW = colW || SW;
    var def = BLOCKS[b.type]; if (!def) return '';
    var sz = b.type === 'banner' || (b.type === 'divider' && +b.size === 0) ? Math.round(clampN(b.size, 0, 600, 0)) : Math.round(clampN(b.size, def.min, def.max, def.def));
    var gap = Math.round(clampN(b.gap, 0, 40, b.type === 'spacer' ? 0 : 4));
    if (b.type === 'spacer') return gapRow(sz);
    var tdOpen = function (style) { return '<tr><td align="' + align + '" style="' + (style || '') + 'padding:0 0 ' + gap + 'px 0;text-align:' + align + ';">'; };
    var textStyle = function () {
      return txt(sz, Math.round(sz * 1.4), colorOf(b.color || def.color, c), (b.bold === undefined ? def.bold : b.bold) ? 'font-weight:bold;' : 'font-weight:normal;') +
        ((b.italic === undefined ? def.italic : b.italic) ? 'font-style:italic;' : '');
    };
    var content = '';
    switch (b.type) {
      case 'logo': content = s.logo_url ? wrapAlign(link(s.website, img(s.logo_url, sz, 0, s.company_name)), align) : ''; return content ? tdOpen() + content + '</td></tr>' : '';
      case 'banner': {
        if (!o.banner || !s.banner_url) return '';
        var bwid = Math.min(sz > 0 ? sz : colW, colW);
        return tdOpen() + wrapAlign(link(s.banner_link, img(s.banner_url, bwid, 0, s.company_name)), align) + '</td></tr>';
      }
      case 'photo': return tdOpen() + wrapAlign(avatar(p, c, sz, b.shape !== 'square'), align) + '</td></tr>';
      case 'name': return p.name ? tdOpen(textStyle()) + esc(p.name) + '</td></tr>' : '';
      case 'title': return p.designation ? tdOpen(textStyle()) + esc(p.designation) + '</td></tr>' : '';
      case 'company': return s.company_name ? tdOpen(textStyle()) + esc(s.company_name) + '</td></tr>' : '';
      case 'locations': content = locations(s); return content ? tdOpen(textStyle()) + esc(content) + '</td></tr>' : '';
      case 'address': content = [s.address, s.address2].filter(Boolean).map(esc).join('<br>'); return content ? tdOpen(textStyle()) + content + '</td></tr>' : '';
      case 'tagline': return s.tagline ? tdOpen(textStyle()) + esc(s.tagline) + '</td></tr>' : '';
      case 'text': return b.text ? tdOpen(textStyle()) + esc(b.text).replace(/\n/g, '<br>') + '</td></tr>' : '';
      case 'mobile': case 'office': case 'email': case 'web': case 'fax': {
        var row = contactRows(p, s).filter(function (x) { return x[0] === def.contact; })[0];
        if (!row) return '';
        var mode = b.labels || LABELMODE || 'text';
        var lab = mode === 'none' ? '' : (mode === 'letters' ? row[1] : row[0] + ':');
        var val = row[2];
        if (b.color) val = val.replace(/color:[^;"]+;/, 'color:' + colorOf(b.color, c) + ';');
        var lc = colorOf(b.labelColor || 'label', c), lh2 = Math.round(sz * 1.5);
        var lw = b.lw == null ? (mode === 'letters' ? 18 : 58) : Math.round(clampN(b.lw, 0, 140, 58));
        if (lab && lw > 0 && align === 'left') {   // fixed label column: values line up when these blocks are stacked
          return '<tr><td style="padding:0 0 ' + gap + 'px 0;">' + tbl('<tr><td width="' + lw + '" valign="top" style="' + txt(sz, lh2, lc, 'width:' + lw + 'px;white-space:nowrap;' + (b.boldLabel ? 'font-weight:bold;' : '')) + '">' + lab + '</td>' +
            '<td valign="top" style="' + txt(sz, lh2, INK) + '">' + val + '</td></tr>') + '</td></tr>';
        }
        return tdOpen(txt(sz, lh2, INK)) + (lab ? '<span style="color:' + lc + ';' + (b.boldLabel ? 'font-weight:bold;' : '') + '">' + lab + '</span>&nbsp;' : '') + val + '</td></tr>';
      }
      case 'contacts': {
        var saved = LABELMODE, savedSep = SEPMODE;
        if (b.labels) LABELMODE = b.labels;
        if (b.sep) SEPMODE = b.sep;
        var lh = Math.round(sz * 1.5);
        if (b.style === 'inline') content = tdOpen(txt(sz, lh, INK)) + contactInline(p, s, 'letters', '|') + '</td></tr>';
        else content = tdOpen() + wrapAlign(contactTable(p, s, 'text').replace(/font-size:12px;line-height:18px/g, 'font-size:' + sz + 'px;line-height:' + lh + 'px'), align) + '</td></tr>';
        LABELMODE = saved; SEPMODE = savedSep;
        return content;
      }
      case 'social': {
        if (!o.social) return '';
        var ic = socialIcons(Object.assign({}, s, { icon_size: String(sz) }), sz);
        return ic ? tdOpen() + wrapAlign(ic, align) + '</td></tr>' : '';
      }
      case 'divider': return tdOpen() + wrapAlign(hline(colorOf(b.color || 'accent', c), sz > 0 ? Math.min(sz, colW) : colW, Math.round(clampN(b.thick, 1, 8, 2))), align) + '</td></tr>';
      case 'button': {
        if (!(s.cta_text && s.cta_link)) return '';
        var bc = colorOf(b.color || 'accent', c);
        var btn = TB + 'border-collapse:separate;"><tr><td bgcolor="' + bc + '" style="background-color:' + bc + ';border-radius:4px;padding:' + Math.round(sz * 0.55) + 'px ' + Math.round(sz * 1.3) + 'px;">' +
          '<a href="' + esc(s.cta_link) + '" target="_blank" style="' + txt(sz, Math.round(sz * 1.35), '#ffffff', 'font-weight:bold;text-decoration:none;display:inline-block;') + '">' + esc(s.cta_text) + '</a></td></tr></table>';
        return tdOpen() + wrapAlign(btn, align) + '</td></tr>';
      }
    }
    return '';
  }
  function designHas(d, type) {
    return (d.rows || []).some(function (r) { return (r.cols || []).some(function (c) { return (c.blocks || []).some(function (b) { return b.type === type; }); }); });
  }
  function designWidth(d) { return Math.round(clampN(d && d.width, 360, MAX_W, SW)); }
  function renderDesign(p, s, o, c, d) {
    var W = designWidth(d);
    var rows = (d.rows || []).map(function (r) {
      var cols = (r.cols || []).filter(Boolean); if (!cols.length) return '';
      var divW = r.divider ? 2 : 0, divGap = r.divider ? 14 : 0, rpad = Math.round(clampN(r.pad, 0, 30, 0));
      var avail = W - 2 * rpad - (cols.length - 1) * (divW + 2 * divGap);
      var sum = cols.reduce(function (a, x) { return a + clampN(x.w, 5, 100, 50); }, 0);
      var cells = cols.map(function (col, i) {
        var w = Math.max(24, Math.floor(avail * clampN(col.w, 5, 100, 50) / sum));
        var al = ['left', 'center', 'right'].indexOf(col.align) >= 0 ? col.align : 'left';
        var va = ['top', 'middle', 'bottom'].indexOf(col.valign) >= 0 ? col.valign : 'top';
        var pad = Math.round(clampN(col.pad, 0, 40, 0)), bg = hex(col.bg, '');
        var inner = (col.blocks || []).map(function (b) { return blockRow(p, s, o, c, b, al, Math.max(20, w - 2 * pad)); }).join('');
        return '<td valign="' + va + '" align="' + al + '" width="' + w + '"' + (bg ? ' bgcolor="' + bg + '"' : '') + ' style="width:' + w + 'px;padding:' + pad + 'px;' + (bg ? 'background-color:' + bg + ';' : '') + '">' +
          (inner ? tbl(inner, Math.max(20, w - 2 * pad)) : '&nbsp;') + '</td>' +
          (i < cols.length - 1 && r.divider ? spacerCell(divGap) + vline(colorOf(r.divColor || 'accent', c), divW) + spacerCell(divGap) : '');
      }).join('');
      var rbg = hex(r.bg, '');
      return '<tr><td style="padding:0 0 ' + Math.round(clampN(r.gap, 0, 40, 8)) + 'px 0;">' +
        tbl('<tr><td' + (rbg ? ' bgcolor="' + rbg + '"' : '') + ' style="padding:' + rpad + 'px;' + (rbg ? 'background-color:' + rbg + ';' : '') + '">' + tbl('<tr>' + cells + '</tr>', W - 2 * rpad) + '</td></tr>', W) + '</td></tr>';
    }).join('');
    return tbl(rows || '<tr><td>&nbsp;</td></tr>', W);
  }
  function findDesign(p, s, extra) {
    var id = String(p.template || '');
    if (id.indexOf('custom:') !== 0) return null;
    var map = (extra && extra.designs) || s._designs || {};
    return map[id.slice(7)] || null;
  }
  var STARTERS = [
    { name: 'Logo left', rows: [{ divider: true, gap: 0, cols: [
      { w: 30, align: 'left', valign: 'middle', blocks: [{ type: 'logo', size: 130 }] },
      { w: 70, blocks: [{ type: 'name' }, { type: 'title', gap: 8 }, { type: 'contacts' }, { type: 'company', gap: 2 }, { type: 'locations', gap: 8 }, { type: 'social' }] }] }] },
    { name: 'Logo on top', rows: [
      { gap: 6, cols: [{ w: 100, blocks: [{ type: 'logo', size: 150 }] }] },
      { gap: 8, cols: [{ w: 100, blocks: [{ type: 'divider', size: 0, thick: 2 }] }] },
      { gap: 0, cols: [{ w: 45, blocks: [{ type: 'name' }, { type: 'title' }, { type: 'company', gap: 8 }, { type: 'social' }] }, { w: 55, blocks: [{ type: 'contacts' }] }] }] },
    { name: 'Photo left', rows: [{ gap: 0, cols: [
      { w: 24, align: 'center', blocks: [{ type: 'photo', size: 88 }] },
      { w: 76, blocks: [{ type: 'name' }, { type: 'title' }, { type: 'company', gap: 8 }, { type: 'divider', size: 40 }, { type: 'contacts', gap: 8 }, { type: 'social' }] }] }] },
    { name: 'Centred card', rows: [{ gap: 0, pad: 16, bg: '#f5f7fa', cols: [{ w: 100, align: 'center', blocks: [
      { type: 'logo', size: 120, gap: 10 }, { type: 'name', size: 18 }, { type: 'title', gap: 8 }, { type: 'divider', size: 60, gap: 8 }, { type: 'contacts', style: 'inline', gap: 10 }, { type: 'social' }] }] }] },
    { name: 'Split contacts', rows: [
      { gap: 8, divider: true, cols: [
        { w: 32, valign: 'middle', blocks: [{ type: 'logo', size: 130 }] },
        { w: 68, valign: 'middle', blocks: [{ type: 'name', size: 18 }, { type: 'title' }, { type: 'company', size: 11 }] }] },
      { gap: 8, cols: [{ w: 100, blocks: [{ type: 'divider', size: 0, thick: 1, color: 'label' }] }] },
      { gap: 6, cols: [
        { w: 50, blocks: [{ type: 'mobile', labels: 'letters' }, { type: 'office', labels: 'letters' }] },
        { w: 50, blocks: [{ type: 'email', labels: 'letters' }, { type: 'web', labels: 'letters' }] }] },
      { gap: 0, cols: [{ w: 100, blocks: [{ type: 'social' }] }] }] },
    { name: 'Blank', rows: [{ gap: 0, cols: [{ w: 100, blocks: [{ type: 'name' }] }] }] }
  ];

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
    if (!s.asset_base_url) return rows ? '<tr><td style="padding:12px 0 0 0;">' + tbl(rows) + '</td></tr>' : '';
    function imgBtn(href, file, w, alt) { return link(href, '<img src="' + assetBase(s) + 'images/buttons/' + file + '" alt="' + alt + '" width="' + w + '" height="37" border="0" style="display:block;width:' + w + 'px;height:37px;border:0;">'); }
    if (s.app_apple) btns.push(imgBtn(s.app_apple, 'apple.png', 119, 'Download on the App Store'));
    if (s.app_google) btns.push(imgBtn(s.app_google, 'google.png', 119, 'Get it on Google Play'));
    if (s.app_amazon) btns.push(imgBtn(s.app_amazon, 'amazon.png', 119, 'Available at Amazon Appstore'));
    if (s.calendar_link) btns.push(imgBtn(s.calendar_link, 'schedule-a-meeting.png', 100, 'Schedule a meeting'));
    var wa = String(s.whatsapp_number || '').replace(/[^\d]/g, '');
    if (wa) btns.push(imgBtn('https://wa.me/' + wa + (s.whatsapp_message ? '?text=' + encodeURIComponent(s.whatsapp_message) : ''), 'whatsapp-contact-me.png', 119, 'Contact me on WhatsApp'));
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
  function applyStyle(s, extra) {
    var c = hex(s.brand_color, '#477ccc');
    INK = hex(s.text_color, '#212121');
    LINK = hex(s.link_color, c);
    MUTE = hex(s.secondary_color, '#6b6b6b');
    FONT = 'font-family:' + (String(s.font_family || '').replace(/["<>;{}]/g, '').trim() || DEFAULT_FONT) + ';';
    LABELMODE = ['text', 'letters', 'none'].indexOf(lower(s.phone_labels)) >= 0 ? lower(s.phone_labels) : '';
    SEPMODE = ['/', '|', 'none'].indexOf(lower(s.separator)) >= 0 ? lower(s.separator) : '';
    ICONDATA = (extra && extra.iconData) || null;
    SW = Math.round(clampN(s.sig_width, 400, MAX_W, MAX_W));
    return c;
  }
  /** One block on its own (used by the Design studio canvas) */
  function blockPreview(p, s, b, align, extra) {
    var c = applyStyle(s, extra), o = options(p);
    var html = blockRow(p, s, o, c, b, align || 'left', (extra && extra.colW) || SW);
    ICONDATA = null;
    return html ? tbl(html, 0, 'width:100%;') : '';
  }

  function render(p, s, extra) {
    var o = options(p);
    var c = applyStyle(s, extra);
    var k = { small: 0.9, large: 1.12 }[lower(s.font_size)] || 1;

    var design = findDesign(p, s, extra);
    var body = design ? renderDesign(p, s, o, c, design) : LAYOUTS[o.template](p, s, o, c);
    var sigW = design ? designWidth(design) : SW;
    var bw = Math.min(parseInt(s.banner_width, 10) || sigW, sigW);   // empty = full width; never wider than the signature
    var showBanner = o.banner && s.banner_url && !(design && designHas(design, 'banner'));
    var banner = showBanner ? '<tr><td style="padding:14px 0 0 0;">' + link(s.banner_link, img(s.banner_url, bw, 0, s.company_name)) + '</td></tr>' : '';
    var w = sigW;
    var disc = o.disclaimer && s.disclaimer
      ? '<tr><td style="padding:14px 0 0 0;">' + tbl('<tr><td style="border-top:1px solid #dddddd;padding:10px 0 0 0;' + txt(9, 12, '#8a8a8a') + '">' + esc(s.disclaimer) + '</td></tr>', w) + '</td></tr>' : '';

    var html = TB + 'width:' + w + 'px;' + FONT + '-webkit-text-size-adjust:none;-ms-text-size-adjust:none;" width="' + w + '">' +
      '<tr><td align="left" style="padding:0;">' + body + '</td></tr>' +
      (o.extras ? extrasBlock(s, c) : '') + banner + (o.badges ? badgesBlock(s) : '') + disc + '</table>';
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
           BLOCKS: BLOCKS, STARTERS: STARTERS, blockPreview: blockPreview, designWidth: designWidth,
           resolveSettings: resolveSettings, render: render, fileHtml: fileHtml, esc: esc, flag: flag, hex: hex,
           iconDesign: iconDesign, iconSetKey: iconSetKey, imgInfo: imgInfo };
})();
