// ============================================
// CA Final Pro Tracker - Service Worker v4.0
// ============================================

const APP_VERSION = '4.0.0';
const CACHE_NAME = `ca-final-pro-tracker-v${APP_VERSION}`;

// Assets to cache on install
const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-72x72.png',
  './icon-96x96.png',
  './icon-128x128.png',
  './icon-144x144.png',
  './icon-192x192.png',
  './icon-512x512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://cdn.jsdelivr.net/npm/chart.js'
];

// Network-first resources (will try network first, then cache)
const NETWORK_FIRST_RESOURCES = [
  'https://fonts.googleapis.com',
  'https://fonts.gstatic.com'
];

// Cache-first resources (will try cache first, then network)
const CACHE_FIRST_RESOURCES = [
  'https://cdnjs.cloudflare.com',
  'https://cdn.jsdelivr.net'
];

// ============================================
// INSTALL EVENT - Cache core assets
// ============================================
self.addEventListener('install', event => {
  console.log('[Service Worker] Installing version:', APP_VERSION);
  
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching core assets');
        return cache.addAll(PRECACHE_ASSETS);
      })
      .then(() => {
        console.log('[Service Worker] Skip waiting on install');
        return self.skipWaiting();
      })
      .catch(error => {
        console.error('[Service Worker] Cache installation failed:', error);
      })
  );
});

// ============================================
// ACTIVATE EVENT - Clean up old caches
// ============================================
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activating version:', APP_VERSION);
  
  event.waitUntil(
    caches.keys()
      .then(cacheNames => {
        return Promise.all(
          cacheNames.map(cacheName => {
            if (cacheName !== CACHE_NAME) {
              console.log('[Service Worker] Deleting old cache:', cacheName);
              return caches.delete(cacheName);
            }
          })
        );
      })
      .then(() => {
        console.log('[Service Worker] Claiming clients');
        return self.clients.claim();
      })
      .then(() => {
        // Schedule notifications after activation
        scheduleNotifications();
        console.log('[Service Worker] Notifications scheduled');
      })
  );
});

// ============================================
// FETCH EVENT - Network strategies
// ============================================
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  
  // Skip non-GET requests
  if (event.request.method !== 'GET') return;
  
  // Skip chrome-extension requests
  if (url.protocol === 'chrome-extension:') return;
  
  // Skip analytics/telemetry
  if (url.hostname.includes('google-analytics') || 
      url.hostname.includes('doubleclick.net')) {
    return;
  }
  
  event.respondWith(
    (async () => {
      // Network-first for API calls and dynamic content
      if (url.pathname.includes('/api/') || url.search.includes('nocache=true')) {
        try {
          return await fetch(event.request);
        } catch (error) {
          const cached = await caches.match(event.request);
          if (cached) return cached;
          throw error;
        }
      }
      
      // Cache-first for CDN resources
      if (CACHE_FIRST_RESOURCES.some(domain => url.href.includes(domain))) {
        const cachedResponse = await caches.match(event.request);
        if (cachedResponse) {
          // Update cache in background
          fetchAndCache(event.request);
          return cachedResponse;
        }
        return fetchAndCache(event.request);
      }
      
      // Network-first for fonts
      if (NETWORK_FIRST_RESOURCES.some(domain => url.href.includes(domain))) {
        try {
          const networkResponse = await fetch(event.request);
          const cache = await caches.open(CACHE_NAME);
          cache.put(event.request, networkResponse.clone());
          return networkResponse;
        } catch (error) {
          const cachedResponse = await caches.match(event.request);
          if (cachedResponse) return cachedResponse;
          throw error;
        }
      }
      
      // Stale-while-revalidate for everything else
      const cachedResponse = await caches.match(event.request);
      const fetchPromise = fetchAndCache(event.request);
      
      return cachedResponse || fetchPromise;
    })()
  );
});

// Helper function to fetch and cache
async function fetchAndCache(request) {
  try {
    const networkResponse = await fetch(request);
    
    // Don't cache non-successful responses
    if (!networkResponse.ok) return networkResponse;
    
    // Don't cache opaque responses (CORS)
    if (networkResponse.type === 'opaque') return networkResponse;
    
    const cache = await caches.open(CACHE_NAME);
    cache.put(request, networkResponse.clone());
    return networkResponse;
  } catch (error) {
    console.error('[Service Worker] Fetch failed:', error);
    throw error;
  }
}

// ============================================
// BACKGROUND SYNC & PERIODIC SYNC
// ============================================
self.addEventListener('sync', event => {
  if (event.tag === 'sync-data') {
    console.log('[Service Worker] Background sync triggered');
    event.waitUntil(syncAppData());
  }
});

self.addEventListener('periodicsync', event => {
  if (event.tag === 'daily-sync' && event.registration) {
    console.log('[Service Worker] Periodic sync triggered');
    event.waitUntil(syncAppData());
  }
});

async function syncAppData() {
  // Here you can implement data sync with a backend if needed
  console.log('[Service Worker] Syncing app data');
  return Promise.resolve();
}

// ============================================
// PUSH NOTIFICATIONS
// ============================================
self.addEventListener('push', event => {
  console.log('[Service Worker] Push received:', event);
  
  let data = { title: 'CA Final Pro Tracker', body: 'New notification' };
  
  if (event.data) {
    try {
      data = event.data.json();
    } catch (e) {
      data.body = event.data.text();
    }
  }
  
  const options = {
    body: data.body,
    icon: './icon-192x192.png',
    badge: './icon-96x96.png',
    vibrate: [200, 100, 200],
    data: {
      url: data.url || './',
      dateOfArrival: Date.now()
    },
    actions: [
      {
        action: 'open',
        title: 'Open App'
      },
      {
        action: 'close',
        title: 'Close'
      }
    ]
  };
  
  event.waitUntil(
    self.registration.showNotification(data.title, options)
  );
});

self.addEventListener('notificationclick', event => {
  console.log('[Service Worker] Notification click:', event);
  
  event.notification.close();
  
  if (event.action === 'close') {
    return;
  }
  
  event.waitUntil(
    clients.matchAll({
      type: 'window',
      includeUncontrolled: true
    }).then(clientList => {
      // Check if there's already a window/tab open
      for (const client of clientList) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          return client.focus();
        }
      }
      // If no window is open, open a new one
      if (clients.openWindow) {
        return clients.openWindow(event.notification.data.url || './');
      }
    })
  );
});

// ============================================
// NOTIFICATION SCHEDULING
// ============================================
function scheduleNotifications() {
  // Clear any existing notifications
  self.registration.getNotifications().then(notifications => {
    notifications.forEach(notification => notification.close());
  });
  
  // Check notification permission
  if (Notification.permission !== 'granted') {
    console.log('[Service Worker] Notifications not granted');
    return;
  }
  
  // Schedule study reminders
  scheduleStudyReminders();
  
  // Schedule periodic notifications every 3 hours during study hours
  schedulePeriodicReminders();
}

function scheduleStudyReminders() {
  const now = new Date();
  
  // Morning reminder at 6 AM
  const morningTime = new Date(now);
  morningTime.setHours(6, 0, 0, 0);
  if (morningTime < now) {
    morningTime.setDate(morningTime.getDate() + 1);
  }
  
  const morningDelay = morningTime.getTime() - Date.now();
  
  setTimeout(() => {
    showNotification(
      '🌅 Morning Study Reminder',
      'Good morning! Plan your study targets for today.',
      'morning-reminder'
    );
    // Reschedule for next day
    scheduleStudyReminders();
  }, morningDelay);
  
  // Evening reminder at 8 PM
  const eveningTime = new Date(now);
  eveningTime.setHours(20, 0, 0, 0);
  if (eveningTime < now) {
    eveningTime.setDate(eveningTime.getDate() + 1);
  }
  
  const eveningDelay = eveningTime.getTime() - Date.now();
  
  setTimeout(() => {
    showNotification(
      '🌙 Evening Progress Check',
      'Evening check! Update your study hours and track progress.',
      'evening-reminder'
    );
  }, eveningDelay);
}

function schedulePeriodicReminders() {
  // Check every 3 hours between 8 AM and 10 PM
  setInterval(() => {
    const now = new Date();
    const hour = now.getHours();
    
    if (hour >= 8 && hour <= 22) {
      const messages = [
        "📚 Time for focused study!",
        "⏰ Stay consistent with your targets!",
        "🎯 Don't forget to log your study hours!",
        "💪 Keep pushing toward your CA goals!",
        "📈 Track your progress for better results!"
      ];
      
      const randomMessage = messages[Math.floor(Math.random() * messages.length)];
      showNotification('CA Study Time', randomMessage, 'periodic-reminder');
    }
  }, 3 * 60 * 60 * 1000); // 3 hours
}

function showNotification(title, body, tag) {
  self.registration.showNotification(title, {
    body: body,
    icon: './icon-192x192.png',
    badge: './icon-96x96.png',
    tag: tag,
    requireInteraction: false,
    silent: false,
    vibrate: [200, 100, 200],
    data: {
      dateOfArrival: Date.now(),
      primaryKey: 1
    }
  }).catch(error => {
    console.error('[Service Worker] Failed to show notification:', error);
  });
}

// ============================================
// MESSAGE HANDLING
// ============================================
self.addEventListener('message', event => {
  console.log('[Service Worker] Message received:', event.data);
  
  if (event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  
  if (event.data.type === 'CLEAR_CACHE') {
    caches.delete(CACHE_NAME).then(() => {
      event.ports[0].postMessage({ success: true });
    });
  }
  
  if (event.data.type === 'GET_CACHE_INFO') {
    caches.open(CACHE_NAME).then(cache => {
      cache.keys().then(keys => {
        event.ports[0].postMessage({ 
          cacheName: CACHE_NAME,
          cacheSize: keys.length,
          version: APP_VERSION 
        });
      });
    });
  }
});

// ============================================
// ERROR HANDLING
// ============================================
self.addEventListener('error', event => {
  console.error('[Service Worker] Error:', event.error);
});

self.addEventListener('unhandledrejection', event => {
  console.error('[Service Worker] Unhandled rejection:', event.reason);
});
