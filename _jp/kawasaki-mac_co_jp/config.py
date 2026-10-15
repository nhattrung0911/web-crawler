# Ultra-performance configuration
PERFORMANCE_CONFIG = {
    # Network
    "HTTP2_ENABLED": True,
    "MAX_CONCURRENT": 100,  # Tăng cho mạng nhanh
    "CONNECTION_TIMEOUT": 10,
    "KEEPALIVE_TIMEOUT": 30,
    
    # Parsing
    "USE_SELECTOLAX": True,
    "CACHE_ENABLED": True,
    "CACHE_TTL": 3600,  # 1 hour
    
    # Processing
    "BATCH_SIZE": 100,
    "CHUNK_SIZE": 8192,
    
    # Memory
    "MAX_MEMORY_MB": 1024,
    "USE_MMAP": True,
    
    # Retry logic
    "MAX_RETRIES": 3,
    "RETRY_BACKOFF": [1, 2, 4],  # Exponential
}