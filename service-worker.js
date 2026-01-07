const CACHE_NAME = 'ca-final-tracker-v4';
const APP_PREFIX = 'CAFINAL_';
const urlsToCache = [
  './',
  './index.html',
  './manifest.json'
];

// Install event - cache resources
self.addEventListener('install', event => {
  console.log('[Service Worker] Install');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching app shell');
        return cache.addAll(urlsToCache);
      })
      .then(() => self.skipWaiting())
  );
});

// Activate event - clean up old caches
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activate');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName.startsWith(APP_PREFIX) && cacheName !== CACHE_NAME) {
            console.log('[Service Worker] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch event - serve from cache or network
self.addEventListener('fetch', event => {
  // Skip non-GET requests
  if (event.request.method !== 'GET') return;
  
  // Skip chrome-extension requests
  if (event.request.url.startsWith('chrome-extension://')) return;
  
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Cache hit - return response
        if (response) {
          return response;
        }
        
        // Clone the request
        const fetchRequest = event.request.clone();
        
        return fetch(fetchRequest).then(response => {
          // Check if we received a valid response
          if (!response || response.status !== 200 || response.type !== 'basic') {
            return response;
          }
          
          // Clone the response
          const responseToCache = response.clone();
          
          caches.open(CACHE_NAME)
            .then(cache => {
              cache.put(event.request, responseToCache);
            });
          
          return response;
        });
      })
  );
});

// Notification click handler
self.addEventListener('notificationclick', event => {
  console.log('[Service Worker] Notification click received.');
  event.notification.close();
  
  event.waitUntil(
    clients.matchAll({type: 'window'})
      .then(clientList => {
        // Focus existing window if available
        for (const client of clientList) {
          if (client.url === self.location.origin && 'focus' in client) {
            return client.focus();
          }
        }
        // Open new window if none exists
        if (clients.openWindow) {
          return clients.openWindow('./');
        }
      })
  );
});

// Notification close handler
self.addEventListener('notificationclose', event => {
  console.log('[Service Worker] Notification closed:', event.notification.tag);
});

// Schedule daily notifications
const scheduleNotifications = () => {
  // Check if notifications are supported
  if (!self.Notification || !self.registration) return;
  
  // Check permission
  if (Notification.permission !== 'granted') return;
  
  // Morning notification at 9:00 AM
  const morningTime = new Date();
  morningTime.setHours(9, 0, 0, 0);
  if (morningTime < new Date()) {
    morningTime.setDate(morningTime.getDate() + 1);
  }
  
  // Evening notification at 8:00 PM
  const eveningTime = new Date();
  eveningTime.setHours(20, 0, 0, 0);
  if (eveningTime < new Date()) {
    eveningTime.setDate(eveningTime.getDate() + 1);
  }
  
  const now = Date.now();
  const morningDelay = morningTime.getTime() - now;
  const eveningDelay = eveningTime.getTime() - now;
  
  // Schedule morning notification
  if (morningDelay > 0) {
    setTimeout(() => {
      self.registration.showNotification('CA Final Pro Tracker', {
        body: '📚 Good morning! Time to plan your study targets for today.',
        icon: './icon-192.png',
        badge: './icon-192.png',
        tag: 'morning-reminder',
        requireInteraction: false,
        vibrate: [200, 100, 200]
      });
      
      // Reschedule for next day
      scheduleNotifications();
    }, morningDelay);
  }
  
  // Schedule evening notification
  if (eveningDelay > 0) {
    setTimeout(() => {
      self.registration.showNotification('CA Final Pro Tracker', {
        body: '📊 Evening check! Update your study hours and track progress.',
        icon: './icon-192.png',
        badge: './icon-192.png',
        tag: 'evening-reminder',
        requireInteraction: false,
        vibrate: [200, 100, 200]
      });
    }, eveningDelay);
  }
};

// Run when service worker is activated
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activated');
  event.waitUntil(scheduleNotifications());
});
