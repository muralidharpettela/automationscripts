import watchdog.events
import watchdog.observers
import time
from update_stock.update_products_data_staging import StagingUpdateProducts


class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self):
        # Set the patterns for PatternMatchingEventHandler
        watchdog.events.PatternMatchingEventHandler.__init__(self, patterns=['*.csv'],
                                                             ignore_directories=True, case_sensitive=False)
    def on_created(self, event):
        print("Watchdog received created event - % s." % event.src_path)

    # Event is created, you can process it now

    def on_modified(self, event):
        print("Watchdog received modified event - % s." % event.src_path)
        staging_products_update = StagingUpdateProducts(event.src_path)
        staging_products_update.process()
    # Event is modified, you can process it now


if __name__ == "__main__":
    src_path = "/Users/muralidharpettela/Documents/"
    event_handler = Handler()
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, path=src_path, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
