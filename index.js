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

function getTimee(time) {
    if (!time) {
        return 0;
    }
    if (/^\d+$/g.test(time)) {
        return parseInt(time);
    }
    const reg = /(\d+)(h|m)/gi;
    const m = 60;
    const h = m * 60;
    const dataTime = time.toString(reg);
    let sum = 0
        n = dataTime
        const hour = n.replace(/(\d+[^\dh])/gi, "").replace(/h+/gi, "");
        const minute = n.replace(/(\d+[^\dm])/gi, "").replace(/m+/gi, "");
        if (hour) {
            sum += (h * Number(hour));
        };
        if (minute) {
            sum += (m * Number(minute));
        }
        return sum;
}

const newXlss = data.map(function(item) {
    // console.log(data);
    const sumDayNotCheckIn = _.get(item, 'Tổng ngày không check in')
    const sumDayNotCheckOut = _.get(item, 'Tổng ngày không check out')
    const sumLateDay = _.get(item, 'Tổng ngày đi trễ')
    const sumLateTime = getTimee(_.get(item, 'Tổng thời gian đi trễ')) / 60
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

