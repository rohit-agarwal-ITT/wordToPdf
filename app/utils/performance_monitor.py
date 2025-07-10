import time
import psutil
import os
import logging
from typing import Dict, Any, Optional, Tuple
from functools import wraps
from contextlib import contextmanager

logger = logging.getLogger(__name__)

class PerformanceMonitor:
    """Monitor application performance and resource usage"""
    
    def __init__(self):
        self.metrics = {}
        self.start_time = None
    
    @contextmanager
    def monitor_operation(self, operation_name: str):
        """Context manager to monitor operation performance"""
        start_time = time.time()
        start_memory = psutil.Process().memory_info().rss
        start_cpu = psutil.cpu_percent()
        
        try:
            yield
        finally:
            end_time = time.time()
            end_memory = psutil.Process().memory_info().rss
            end_cpu = psutil.cpu_percent()
            
            duration = end_time - start_time
            memory_delta = end_memory - start_memory
            
            self.record_metric(operation_name, {
                'duration': duration,
                'memory_delta': memory_delta,
                'peak_memory': max(start_memory, end_memory),
                'avg_cpu': (start_cpu + end_cpu) / 2
            })
            
            logger.info(f"Operation '{operation_name}' completed in {duration:.2f}s, "
                       f"memory delta: {memory_delta / 1024 / 1024:.2f}MB")
    
    def record_metric(self, operation: str, metrics: Dict[str, Any]):
        """Record performance metrics for an operation"""
        if operation not in self.metrics:
            self.metrics[operation] = []
        
        self.metrics[operation].append({
            'timestamp': time.time(),
            **metrics
        })
    
    def get_system_info(self) -> Dict[str, Any]:
        """Get current system resource information"""
        try:
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            cpu_count = psutil.cpu_count()
            
            return {
                'cpu_percent': psutil.cpu_percent(interval=1),
                'memory_total': memory.total,
                'memory_available': memory.available,
                'memory_percent': memory.percent,
                'disk_total': disk.total,
                'disk_free': disk.free,
                'disk_percent': (disk.used / disk.total) * 100,
                'cpu_count': cpu_count
            }
        except Exception as e:
            logger.error(f"Error getting system info: {e}")
            return {}
    
    def check_system_health(self) -> Tuple[bool, str]:
        """Check if system has enough resources for conversion"""
        try:
            system_info = self.get_system_info()
            
            # Check memory (need at least 500MB free)
            memory_available_mb = system_info.get('memory_available', 0) / 1024 / 1024
            if memory_available_mb < 500:
                return False, f"Insufficient memory. Only {memory_available_mb:.0f}MB available, need at least 500MB"
            
            # Check disk space (need at least 1GB free)
            disk_free_gb = system_info.get('disk_free', 0) / 1024 / 1024 / 1024
            if disk_free_gb < 1:
                return False, f"Insufficient disk space. Only {disk_free_gb:.1f}GB available, need at least 1GB"
            
            # Check CPU usage (should be below 90%)
            cpu_percent = system_info.get('cpu_percent', 0)
            if cpu_percent > 90:
                return False, f"High CPU usage: {cpu_percent:.1f}%. Please try again later"
            
            return True, "System healthy"
            
        except Exception as e:
            logger.error(f"Error checking system health: {e}")
            return False, f"Unable to check system health: {str(e)}"
    
    def get_performance_summary(self) -> Dict[str, Any]:
        """Get performance summary for all recorded operations"""
        summary = {}
        
        for operation, metrics_list in self.metrics.items():
            if not metrics_list:
                continue
            
            durations = [m['duration'] for m in metrics_list]
            memory_deltas = [m['memory_delta'] for m in metrics_list]
            
            summary[operation] = {
                'count': len(metrics_list),
                'avg_duration': sum(durations) / len(durations),
                'min_duration': min(durations),
                'max_duration': max(durations),
                'avg_memory_delta': sum(memory_deltas) / len(memory_deltas),
                'total_memory_delta': sum(memory_deltas)
            }
        
        return summary
    
    def cleanup_old_metrics(self, max_age_hours: int = 24):
        """Clean up metrics older than specified hours"""
        current_time = time.time()
        cutoff_time = current_time - (max_age_hours * 3600)
        
        for operation in list(self.metrics.keys()):
            self.metrics[operation] = [
                m for m in self.metrics[operation]
                if m['timestamp'] > cutoff_time
            ]
            
            # Remove empty operations
            if not self.metrics[operation]:
                del self.metrics[operation]

# Global performance monitor instance
performance_monitor = PerformanceMonitor()

def monitor_performance(operation_name: str):
    """Decorator to monitor function performance"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            with performance_monitor.monitor_operation(operation_name):
                return func(*args, **kwargs)
        return wrapper
    return decorator 