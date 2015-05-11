/*global sfss, Logger */
Logger.log('entering file r_monitor');
try {
    if (sfss.r.MONITOR.b) {
        sfss.r.wrap();
    }
} catch (e) {
    Logger.log(sfss.lg.catchToString(e));
}
Logger.log('leaving file r_monitor');