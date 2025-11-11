import { Client } from '@microsoft/microsoft-graph-client';
import { DeviceCodeCredential } from '@azure/identity';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const tokenFilePath = path.join(__dirname, '../../.access-token.txt');

export const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
export const scopes = [
  'Notes.Read.All',
  'Notes.ReadWrite.All',
  'Sites.Read.All',
  'Group.Read.All',
  'User.Read'
];

export let accessToken = null;
export let graphClient = null;

// Cache system to improve performance
const cache = new Map();

// Cache TTL (time to live) in milliseconds
const CACHE_TTL = {
  groups: 60 * 60 * 1000,       // 1 hour (rarely changes)
  notebooks: 30 * 60 * 1000,    // 30 minutes (rarely changes)
  sections: 30 * 60 * 1000,     // 30 minutes (rarely changes)
  sectionGroups: 30 * 60 * 1000,// 30 minutes (rarely changes)
  pages: 10 * 60 * 1000,        // 10 minutes (changes more often)
  pageContent: 5 * 60 * 1000,   // 5 minutes (can change frequently)
  default: 15 * 60 * 1000       // 15 minutes default
};

// Determine cache TTL based on API path
function getCacheTTL(apiPath) {
  if (apiPath.includes('/memberOf')) return CACHE_TTL.groups;
  if (apiPath.includes('/notebooks') && !apiPath.includes('/sections')) return CACHE_TTL.notebooks;
  if (apiPath.includes('/sectionGroups')) return CACHE_TTL.sectionGroups;
  if (apiPath.includes('/sections') && !apiPath.includes('/pages')) return CACHE_TTL.sections;
  if (apiPath.includes('/pages/') && apiPath.includes('/content')) return CACHE_TTL.pageContent;
  if (apiPath.includes('/pages')) return CACHE_TTL.pages;
  return CACHE_TTL.default;
}

// Get cached data if still valid
function getCachedData(cacheKey) {
  const cached = cache.get(cacheKey);
  if (!cached) return null;

  const now = Date.now();
  if (now - cached.timestamp > cached.ttl) {
    cache.delete(cacheKey);
    return null;
  }

  return cached.data;
}

// Store data in cache
function setCachedData(cacheKey, data, ttl) {
  cache.set(cacheKey, {
    data: data,
    timestamp: Date.now(),
    ttl: ttl
  });
}

// Clear all cache (useful for testing or after updates)
export function clearCache() {
  cache.clear();
  console.error('Cache cleared');
}

// Get cache stats
export function getCacheStats() {
  return {
    entries: cache.size,
    keys: Array.from(cache.keys())
  };
}

// Wrap Graph API call with caching
export async function cachedApiCall(apiPath, options = {}) {
  await ensureGraphClient();

  // Create cache key from path and options
  const cacheKey = `${apiPath}${JSON.stringify(options)}`;

  // Check cache first
  const cachedData = getCachedData(cacheKey);
  if (cachedData) {
    console.error(`Cache hit: ${apiPath}`);
    return cachedData;
  }

  // Make API call
  console.error(`Cache miss: ${apiPath}`);
  let apiRequest = graphClient.api(apiPath);

  // Apply options (select, top, filter, etc.)
  if (options.select) apiRequest = apiRequest.select(options.select);
  if (options.top) apiRequest = apiRequest.top(options.top);
  if (options.filter) apiRequest = apiRequest.filter(options.filter);
  if (options.orderby) apiRequest = apiRequest.orderby(options.orderby);

  const data = await apiRequest.get();

  // Cache the result
  const ttl = getCacheTTL(apiPath);
  setCachedData(cacheKey, data, ttl);

  return data;
}

// Try to read the stored access token on module load
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
}

if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

export async function ensureGraphClient() {
  if (!graphClient) {
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
        try {
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.token;
        } catch (parseError) {
          accessToken = tokenData;
        }
      }
    } catch (error) {
      console.error("Error reading token file:", error);
    }

    if (!accessToken) {
      throw new Error("Access token not found. Please authenticate first.");
    }

    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
  }
  return graphClient;
}

export async function createGraphClient() {
  // If we have a token, test if it's still valid
  if (accessToken) {
    try {
      const testClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return accessToken;
          }
        }
      });

      // Try a simple API call to validate the token
      await testClient.api('/me').get();

      // Token is valid
      graphClient = testClient;
      return { type: 'token', client: graphClient };
    } catch (error) {
      // Token is invalid or expired
      console.error('Existing token is invalid or expired, starting new authentication...');
      accessToken = null;
      graphClient = null;

      // Delete the invalid token file
      try {
        if (fs.existsSync(tokenFilePath)) {
          fs.unlinkSync(tokenFilePath);
        }
      } catch (deleteError) {
        console.error('Error deleting token file:', deleteError.message);
      }
    }
  }

  // No token or invalid token - start device code flow
  const credential = new DeviceCodeCredential({
    clientId: clientId,
    userPromptCallback: (info) => {
      console.error('\n' + info.message);
    }
  });

  try {
    const tokenResponse = await credential.getToken(scopes);

    accessToken = tokenResponse.token;
    fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));

    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });

    return { type: 'device_code', client: graphClient };
  } catch (error) {
    console.error('Authentication error:', error);
    throw new Error(`Authentication failed: ${error.message}`);
  }
}

export { tokenFilePath };
