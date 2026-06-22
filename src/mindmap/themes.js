/**
 * Color themes for the Mind Map Studio add-in.
 * Each theme defines colors for up to 6 depth levels + root.
 * level[0] = root's direct children, level[1] = grandchildren, etc.
 */
const THEMES = {
  classic: {
    id: 'classic',
    name: 'Classic',
    canvas: '#f7f9fc',
    root: { bg: '#2c3e50', text: '#ffffff', border: '#1a252f' },
    levels: [
      { bg: '#2980b9', text: '#ffffff', border: '#1a6fa3' },
      { bg: '#27ae60', text: '#ffffff', border: '#1e8449' },
      { bg: '#8e44ad', text: '#ffffff', border: '#7d3c98' },
      { bg: '#e67e22', text: '#ffffff', border: '#ca6f1e' },
      { bg: '#16a085', text: '#ffffff', border: '#138d75' },
      { bg: '#c0392b', text: '#ffffff', border: '#a93226' }
    ],
    edge: '#95a5a6'
  },

  pastel: {
    id: 'pastel',
    name: 'Pastel Dream',
    canvas: '#fefefe',
    root: { bg: '#b8d4e8', text: '#34495e', border: '#8ab4d4' },
    levels: [
      { bg: '#aed6f1', text: '#2c3e50', border: '#85c1e9' },
      { bg: '#a9dfbf', text: '#1e8449', border: '#7dcea0' },
      { bg: '#f9e79f', text: '#7d6608', border: '#f7dc6f' },
      { bg: '#f5cba7', text: '#935116', border: '#f0a57b' },
      { bg: '#d2b4de', text: '#6c3483', border: '#c39bd3' },
      { bg: '#fadbd8', text: '#922b21', border: '#f5b7b1' }
    ],
    edge: '#d5d8dc'
  },

  pink: {
    id: 'pink',
    name: 'Pink Dream',
    canvas: '#fff5f8',
    root: { bg: '#c2185b', text: '#ffffff', border: '#ad1457' },
    levels: [
      { bg: '#e91e8c', text: '#ffffff', border: '#c2177a' },
      { bg: '#ff69b4', text: '#ffffff', border: '#e0559d' },
      { bg: '#ffb6c1', text: '#7b2050', border: '#ff8fa3' },
      { bg: '#ff85a1', text: '#ffffff', border: '#e0607f' },
      { bg: '#ffc0cb', text: '#7b2050', border: '#ffa0b0' },
      { bg: '#ff1493', text: '#ffffff', border: '#d40079' }
    ],
    edge: '#f48fb1'
  },

  ocean: {
    id: 'ocean',
    name: 'Deep Ocean',
    canvas: '#f0f8ff',
    root: { bg: '#03045e', text: '#ffffff', border: '#020337' },
    levels: [
      { bg: '#0077b6', text: '#ffffff', border: '#005f92' },
      { bg: '#0096c7', text: '#ffffff', border: '#007aad' },
      { bg: '#00b4d8', text: '#ffffff', border: '#009ab8' },
      { bg: '#48cae4', text: '#023e5e', border: '#28b6d4' },
      { bg: '#90e0ef', text: '#023e5e', border: '#5ccfe4' },
      { bg: '#caf0f8', text: '#023e5e', border: '#9ad8e4' }
    ],
    edge: '#90e0ef'
  },

  forest: {
    id: 'forest',
    name: 'Enchanted Forest',
    canvas: '#f1faee',
    root: { bg: '#1b4332', text: '#ffffff', border: '#081c15' },
    levels: [
      { bg: '#2d6a4f', text: '#ffffff', border: '#1f4f3a' },
      { bg: '#40916c', text: '#ffffff', border: '#2f7153' },
      { bg: '#52b788', text: '#ffffff', border: '#3d9a6e' },
      { bg: '#74c69d', text: '#1b4332', border: '#52b084' },
      { bg: '#95d5b2', text: '#1b4332', border: '#6ec29a' },
      { bg: '#d8f3dc', text: '#1b4332', border: '#a8debb' }
    ],
    edge: '#95d5b2'
  }
};

/**
 * Returns the theme colors for a given node depth.
 * depth 0 = root, depth 1 = direct children, etc.
 */
function getThemeColors(themeId, depth) {
  const theme = THEMES[themeId] || THEMES.classic;
  if (depth === 0) return { ...theme.root };
  const levelIdx = Math.min(depth - 1, theme.levels.length - 1);
  return { ...theme.levels[levelIdx] };
}

if (typeof module !== 'undefined') module.exports = { THEMES, getThemeColors };
