const xlss = require('xlsx')
const _ = require('lodash'); // 
const file = xlss.readFile("./abc.xlsx", { type: 'binary' });

const POINT = {
    CheckLate: 100,
    LATE_PER_MINUTE: 1,
    NOTE_CHECK_IN: 500,
    NOT_CHECK_OUT: 50,
};

const data = (xlss.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]));

const formatDisplay = {
    PRIMARY: 'primary'
};

const getMS = {
    h: {
        [formatDisplay.PRIMARY]: 'H'
    },
    m: {
        [formatDisplay.PRIMARY]: 'M'
    }
}

function getTimeConst({ m, h }, format) {
    return {
        [getMS.h[format]]: h,
        [getMS.m[format]]: m
    }
}

function getTimeMS(time) {
    if (!time) {
        return 0; // không phải time trả về 0
    }
    if (/^\d+$/g.test(time)) {
        return parseInt(time);
    }
    const format = formatDisplay.PRIMARY;
    const reg = /(\d+)(\w)/g;
    const m = 1000 * 60; // 1p = 60000 ms
    const h = m * 60; // 1h = 3600000 ms
    const timeConst = getTimeConst({ m, h }, format);

    return _.reduce(time.match(reg), function(sum, n) {
        const key = n.slice(-1); // 1 : H ; 2 : M
        let value = parseInt(n.slice(0, -1)); // 16  // 59
        value = /^\d+$/g.test(value) ? parseInt(n.slice(0, -1)) : 0; // 16  // 59
        if (!timeConst[key] || !/^\d+$/g.test(value)) {
            return sum + 0;
        }
        return sum + (timeConst[key] * value); // 0 + 3600000 *16 = 57600000     ///  57600000 + (60000 * 59)  = 61140000 / 60000 = 1019
    }, 0);
}

const newXlss = data.map(function(item) {

    // console.log(data);
    const sumDayNotCheckIn = _.get(item, 'Tổng ngày không check in')
    const sumDayNotCheckOut = _.get(item, 'Tổng ngày không check out')
    const sumLateDay = _.get(item, 'Tổng ngày đi trễ')
    const sumLateTime = getTimeMS(_.get(item, 'Tổng thời gian đi trễ')) / 60000
    console.log(sumLateTime, ' sumlateTime')
    const sum = sumDayNotCheckIn * POINT.NOTE_CHECK_IN + sumDayNotCheckOut * POINT.NOT_CHECK_OUT + sumLateDay * POINT.CheckLate + sumLateTime * POINT.LATE_PER_MINUTE;

    return {
        name: item['Họ và tên'],
        email: item['Email'],
        sum
    }

});

_.forEach(_.orderBy(newXlss, ['sum'], ['desc']), (item, index) => {
    console.log(index, item);
});