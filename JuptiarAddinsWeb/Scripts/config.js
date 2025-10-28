/**
 * Configuration file for Jupiter Addins
 * Contains server settings and other configuration options
 */

/**
 * Modern DOM Utility Library - Replaces jQuery functionality
 * Provides a lightweight, modern alternative to jQuery
 */
class DOMUtils {
    /**
     * Select single element (replaces $())
     * @param {string} selector - CSS selector
     * @param {Element} context - Context element (optional)
     * @returns {Element|null}
     */
    static select(selector, context = document) {
        return context.querySelector(selector);
    }

    /**
     * Select multiple elements (replaces $())
     * @param {string} selector - CSS selector
     * @param {Element} context - Context element (optional)
     * @returns {NodeList}
     */
    static selectAll(selector, context = document) {
        return context.querySelectorAll(selector);
    }

    /**
     * Add event listener (replaces .on())
     * @param {Element|string} element - Element or selector
     * @param {string} event - Event type
     * @param {Function} handler - Event handler
     * @param {boolean} useCapture - Use capture phase
     */
    static on(element, event, handler, useCapture = false) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.addEventListener(event, handler, useCapture);
        }
    }

    /**
     * Add delegated event listener (replaces $(document).on())
     * @param {Element|string} container - Container element or selector
     * @param {string} event - Event type
     * @param {string} selector - Target selector
     * @param {Function} handler - Event handler
     */
    static delegate(container, event, selector, handler) {
        const containerEl = typeof container === 'string' ? this.select(container) : container;
        if (containerEl) {
            containerEl.addEventListener(event, (e) => {
                const matchedElement = e.target.closest(selector);
                if (matchedElement) {
                    // Set currentTarget to the matched element for compatibility
                    Object.defineProperty(e, 'currentTarget', {
                        value: matchedElement,
                        configurable: true
                    });
                    handler(e);
                }
            });
        }
    }

    /**
     * Remove event listener
     * @param {Element|string} element - Element or selector
     * @param {string} event - Event type
     * @param {Function} handler - Event handler
     */
    static off(element, event, handler) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.removeEventListener(event, handler);
        }
    }

    /**
     * Set or get text content (replaces .text())
     * @param {Element|string} element - Element or selector
     * @param {string} text - Text to set (optional)
     * @returns {string|void}
     */
    static text(element, text) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (text !== undefined) {
            el.textContent = text;
        } else {
            return el.textContent;
        }
    }

    /**
     * Set or get HTML content (replaces .html())
     * @param {Element|string} element - Element or selector
     * @param {string} html - HTML to set (optional)
     * @returns {string|void}
     */
    static html(element, html) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (html !== undefined) {
            el.innerHTML = html;
        } else {
            return el.innerHTML;
        }
    }

    /**
     * Set or get input value (replaces .val())
     * @param {Element|string} element - Element or selector
     * @param {string} value - Value to set (optional)
     * @returns {string|void}
     */
    static val(element, value) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (value !== undefined) {
            el.value = value;
        } else {
            return el.value;
        }
    }

    /**
     * Show element (replaces .show())
     * @param {Element|string} element - Element or selector
     * @param {string} display - Display value (default: 'block')
     */
    static show(element, display = 'block') {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.style.display = display;
        }
    }

    /**
     * Hide element (replaces .hide())
     * @param {Element|string} element - Element or selector
     */
    static hide(element) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.style.display = 'none';
        }
    }

    /**
     * Toggle element visibility
     * @param {Element|string} element - Element or selector
     * @param {boolean|string} condition - Boolean condition or display value when showing (default: 'block')
     */
    static toggle(element, condition = null) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (typeof condition === 'boolean') {
            // jQuery-style toggle with condition
            if (condition) {
                this.show(el);
            } else {
                this.hide(el);
            }
        } else {
            // Standard toggle behavior
            const display = typeof condition === 'string' ? condition : 'block';
            if (el.style.display === 'none' || getComputedStyle(el).display === 'none') {
                this.show(el, display);
            } else {
                this.hide(el);
            }
        }
    }

    /**
     * Add CSS class (replaces .addClass())
     * @param {Element|string} element - Element or selector
     * @param {string} className - Class name to add
     */
    static addClass(element, className) {
        if (typeof element === 'string') {
            // If it's a selector, apply to all matching elements
            const elements = this.selectAll(element);
            elements.forEach(el => el.classList.add(className));
        } else if (element) {
            // If it's a single element
            element.classList.add(className);
        }
    }

    /**
     * Remove CSS class (replaces .removeClass())
     * @param {Element|string} element - Element or selector
     * @param {string} className - Class name to remove
     */
    static removeClass(element, className) {
        if (typeof element === 'string') {
            // If it's a selector, apply to all matching elements
            const elements = this.selectAll(element);
            elements.forEach(el => el.classList.remove(className));
        } else if (element) {
            // If it's a single element
            element.classList.remove(className);
        }
    }

    /**
     * Toggle CSS class
     * @param {Element|string} element - Element or selector
     * @param {string} className - Class name to toggle
     */
    static toggleClass(element, className) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.classList.toggle(className);
        }
    }

    /**
     * Check if element has CSS class
     * @param {Element|string} element - Element or selector
     * @param {string} className - Class name to check
     * @returns {boolean}
     */
    static hasClass(element, className) {
        const el = typeof element === 'string' ? this.select(element) : element;
        return el ? el.classList.contains(className) : false;
    }

    /**
     * Set or get attribute (replaces .attr())
     * @param {Element|string} element - Element or selector
     * @param {string} name - Attribute name
     * @param {string} value - Attribute value (optional)
     * @returns {string|void}
     */
    static attr(element, name, value) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (value !== undefined) {
            el.setAttribute(name, value);
        } else {
            return el.getAttribute(name);
        }
    }

    /**
     * Remove attribute
     * @param {Element|string} element - Element or selector
     * @param {string} name - Attribute name
     */
    static removeAttr(element, name) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.removeAttribute(name);
        }
    }

    /**
     * Set or get property (replaces .prop())
     * @param {Element|string} element - Element or selector
     * @param {string} name - Property name
     * @param {any} value - Property value (optional)
     * @returns {any}
     */
    static prop(element, name, value) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (value !== undefined) {
            el[name] = value;
        } else {
            return el[name];
        }
    }

    /**
     * Set CSS styles
     * @param {Element|string} element - Element or selector
     * @param {Object|string} styles - Styles object or property name
     * @param {string} value - Property value (if styles is string)
     */
    static css(element, styles, value) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        if (typeof styles === 'string') {
            if (value !== undefined) {
                el.style[styles] = value;
            } else {
                return getComputedStyle(el)[styles];
            }
        } else if (typeof styles === 'object') {
            Object.assign(el.style, styles);
        }
    }

    /**
     * Get element data attribute (replaces .data())
     * @param {Element|string} element - Element or selector
     * @param {string} key - Data key
     * @param {any} value - Data value (optional)
     * @returns {any}
     */
    static data(element, key, value) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return;

        // Convert kebab-case to camelCase for dataset API
        const camelKey = key.replace(/-([a-z])/g, (match, letter) => letter.toUpperCase());

        if (value !== undefined) {
            el.dataset[camelKey] = value;
        } else {
            return el.dataset[camelKey];
        }
    }

    /**
     * Empty element content (replaces .empty())
     * @param {Element|string} element - Element or selector
     */
    static empty(element) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.innerHTML = '';
        }
    }

    /**
     * Append content to element (replaces .append())
     * @param {Element|string} element - Element or selector
     * @param {string|Element} content - Content to append
     */
    static append(element, content) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            if (typeof content === 'string') {
                el.insertAdjacentHTML('beforeend', content);
            } else {
                el.appendChild(content);
            }
        }
    }

    /**
     * Prepend content to element (replaces .prepend())
     * @param {Element|string} element - Element or selector
     * @param {string|Element} content - Content to prepend
     */
    static prepend(element, content) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            if (typeof content === 'string') {
                el.insertAdjacentHTML('afterbegin', content);
            } else {
                el.insertBefore(content, el.firstChild);
            }
        }
    }

    /**
     * Remove element (replaces .remove())
     * @param {Element|string} element - Element or selector
     */
    static remove(element) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el && el.parentNode) {
            el.parentNode.removeChild(el);
        }
    }

    /**
     * Get element's parent (replaces .parent())
     * @param {Element|string} element - Element or selector
     * @returns {Element|null}
     */
    static parent(element) {
        const el = typeof element === 'string' ? this.select(element) : element;
        return el ? el.parentElement : null;
    }

    /**
     * Find closest ancestor matching selector (replaces .closest())
     * @param {Element|string} element - Element or selector
     * @param {string} selector - Selector to match
     * @returns {Element|null}
     */
    static closest(element, selector) {
        const el = typeof element === 'string' ? this.select(element) : element;
        return el ? el.closest(selector) : null;
    }

    /**
     * Find elements within element (replaces .find())
     * @param {Element|string} element - Element or selector
     * @param {string} selector - Selector to find
     * @returns {NodeList}
     */
    static find(element, selector) {
        const el = typeof element === 'string' ? this.select(element) : element;
        return el ? el.querySelectorAll(selector) : [];
    }

    /**
     * Focus element
     * @param {Element|string} element - Element or selector
     */
    static focus(element) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            el.focus();
        }
    }

    /**
     * Trigger custom event
     * @param {Element|string} element - Element or selector
     * @param {string} eventType - Event type
     * @param {Object} detail - Event detail data
     */
    static trigger(element, eventType, detail = {}) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (el) {
            const event = new CustomEvent(eventType, { detail });
            el.dispatchEvent(event);
        }
    }

    /**
     * Smooth slide up animation (replaces jQuery .slideUp())
     * @param {Element|string} element - Element or selector
     * @param {number} duration - Animation duration in ms
     * @returns {Promise}
     */
    static slideUp(element, duration = 200) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return Promise.resolve();

        return new Promise(resolve => {
            const height = el.offsetHeight;
            el.style.transition = `height ${duration}ms ease-out`;
            el.style.overflow = 'hidden';
            el.style.height = height + 'px';

            requestAnimationFrame(() => {
                el.style.height = '0px';
                setTimeout(() => {
                    el.style.display = 'none';
                    el.style.transition = '';
                    el.style.height = '';
                    el.style.overflow = '';
                    resolve();
                }, duration);
            });
        });
    }

    /**
     * Smooth slide down animation (replaces jQuery .slideDown())
     * @param {Element|string} element - Element or selector
     * @param {number} duration - Animation duration in ms
     * @returns {Promise}
     */
    static slideDown(element, duration = 200) {
        const el = typeof element === 'string' ? this.select(element) : element;
        if (!el) return Promise.resolve();

        return new Promise(resolve => {
            el.style.display = 'block';
            const height = el.scrollHeight;
            el.style.height = '0px';
            el.style.overflow = 'hidden';
            el.style.transition = `height ${duration}ms ease-out`;

            requestAnimationFrame(() => {
                el.style.height = height + 'px';
                setTimeout(() => {
                    el.style.transition = '';
                    el.style.height = '';
                    el.style.overflow = '';
                    resolve();
                }, duration);
            });
        });
    }

    /**
     * Wait for DOM to be ready (replaces $(document).ready())
     * @param {Function} callback - Callback function
     */
    static ready(callback) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', callback);
        } else {
            callback();
        }
    }
}

// Create global shorthand similar to jQuery
window.$ = DOMUtils;
window.DOM = DOMUtils;

window.JupiterConfig = {
    // Server Configuration
    server: {
        // Base URL of the Jupiter DMS API server
        baseUrl: 'https://localhost:7001',

        // API endpoint path
        apiEndpoint: '/api',

        // Request timeout in milliseconds
        timeout: 30000,

        // Health check endpoint
        healthEndpoint: '/health'
    },
    
    // Authentication Configuration
    auth: {
        // Whether to remember credentials by default
        rememberCredentials: false,
        
        // Whether to enable auto-login
        autoLogin: false,
        
        // Token refresh threshold (minutes before expiry)
        refreshThreshold: 5
    },
    
    // UI Configuration
    ui: {
        // Default page size for document listings
        defaultPageSize: 50,
        
        // Maximum file size for uploads (in bytes)
        maxFileSize: 100 * 1024 * 1024, // 100MB
        
        // Supported file types for upload
        supportedFileTypes: [
            '.doc', '.docx', '.pdf', '.txt', '.rtf',
            '.xls', '.xlsx', '.ppt', '.pptx',
            '.jpg', '.jpeg', '.png', '.gif', '.bmp',
            '.zip', '.rar', '.7z'
        ],
        
        // Auto-refresh interval for document lists (milliseconds)
        autoRefreshInterval: 30000
    },
    
    // Feature Flags
    features: {
        // Enable document search functionality
        enableSearch: true,
        
        // Enable document versioning
        enableVersioning: true,
        
        // Enable document check-in/check-out
        enableCheckInOut: true,
        
        // Enable document metadata editing
        enableMetadataEdit: true,
        
        // Enable folder management
        enableFolderManagement: true
    },
    
    // Logging Configuration
    logging: {
        // Log level: 'debug', 'info', 'warn', 'error'
        level: 'info',
        
        // Whether to log to console
        console: true,
        
        // Whether to send logs to server
        remote: false
    },
    
    // Development/Debug Settings
    debug: {
        // Enable debug mode
        enabled: false,
        
        // Show detailed error messages
        verboseErrors: false,
        
        // Enable API request/response logging
        logApiCalls: false
    }
};

/**
 * Get configuration value by path
 * @param {string} path - Dot-separated path to config value (e.g., 'server.baseUrl')
 * @param {*} defaultValue - Default value if path not found
 * @returns {*} Configuration value
 */
window.JupiterConfig.get = function(path, defaultValue = null) {
    const keys = path.split('.');
    let current = this;
    
    for (const key of keys) {
        if (current && typeof current === 'object' && key in current) {
            current = current[key];
        } else {
            return defaultValue;
        }
    }
    
    return current;
};

/**
 * Set configuration value by path
 * @param {string} path - Dot-separated path to config value
 * @param {*} value - Value to set
 */
window.JupiterConfig.set = function(path, value) {
    const keys = path.split('.');
    const lastKey = keys.pop();
    let current = this;
    
    for (const key of keys) {
        if (!(key in current) || typeof current[key] !== 'object') {
            current[key] = {};
        }
        current = current[key];
    }
    
    current[lastKey] = value;
};

/**
 * Initialize configuration with environment-specific overrides
 */
window.JupiterConfig.init = function() {
    // Check for environment-specific configuration
    const hostname = window.location.hostname;
    
    // Development environment
    if (hostname === 'localhost' || hostname === '127.0.0.1') {
        this.set('server.baseUrl', 'https://localhost:7001');
        this.set('debug.enabled', true);
        this.set('logging.level', 'debug');
    }
    
    // Production environment adjustments can be added here
    // Example:
    // if (hostname === 'your-production-domain.com') {
    //     this.set('server.baseUrl', 'https://api.your-production-domain.com');
    //     this.set('debug.enabled', false);
    //     this.set('logging.level', 'warn');
    // }
    
    console.log('Jupiter Config initialized:', this);
};

// Auto-initialize when script loads
document.addEventListener('DOMContentLoaded', function() {
    window.JupiterConfig.init();
});
