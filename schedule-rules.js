(function (root) {
  'use strict';
  var DAY = 86400000;
  function utc(iso) { return Date.parse(iso + 'T00:00:00Z'); }
  function shift(iso, days) { return new Date(utc(iso) + days * DAY).toISOString().slice(0, 10); }
  function dates(from, until, step, eligible) {
    if (!Number.isFinite(utc(from)) || !Number.isFinite(utc(until)) || until < from) throw new Error('Проверьте даты повторения.');
    if (utc(until) - utc(from) > 370 * DAY) throw new Error('Повторение ограничено одним учебным годом.');
    var result = [];
    for (var d = from; d <= until; d = shift(d, step === 14 ? 14 : 7)) {
      if (d === from || !eligible || eligible(d)) result.push(d);
    }
    return result;
  }
  function expand(base, dates, anchor, lookup) {
    var out = [];
    dates.forEach(function (day) {
      var delta = Math.round((utc(day) - utc(anchor)) / DAY);
      base.forEach(function (record) {
        var next = Object.assign({}, record, { date: shift(record.date, delta) });
        if (delta) {
          next.id = '';
          if (next.link) next.link = next.link + '-' + day;
          if (lookup) {
            var actual = lookup(next) || {};
            next.id = actual.id || '';
            next.was = actual.was || '';
          }
        }
        out.push(next);
      });
    });
    return out;
  }
  function inPeriod(day, start, half2, end, half) {
    return day >= start && day < end && (half === 0 || (half === 1 ? day < half2 : day >= half2));
  }
  var api = { dates: dates, expand: expand, shift: shift, inPeriod: inPeriod };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.RaspRules = api;
})(typeof window === 'undefined' ? this : window);

