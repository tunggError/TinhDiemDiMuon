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

// const getMS = {
//     h: 'H',
//     m: 'M'
// }

// function getTimeConst({ m, h }) {
//     return {
//         [getMS.h]: h,
//         [getMS.m]: m
//     }
// }

// function getTimeMS(time) {
//     if (!time) {
//         return 0;
//     }
//     if (/^\d+$/g.test(time)) {
//         return parseInt(time);
//     }
//     const reg = /(\d+)(\w)/g;
//     const m = 60;
//     const h = m * 60;
//     const timeConst = getTimeConst({ m, h });

//     return _.reduce(time.match(reg), function(sum, n) {
//         const key = n.slice(-1);
//         let value = parseInt(n.slice(0, -1));
//         value = /^\d+$/g.test(value) ? parseInt(n.slice(0, -1)) : 0;
//         if (!timeConst[key] || !/^\d+$/g.test(value)) {

//             return sum + 0;
//         }
//         return sum + (timeConst[key] * value);
//     }, 0);
// }




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


function getTimee(time, sum) {
    if (!time) {
        return 0;
    }
    if (/^\d+$/g.test(time)) {
        return parseInt(time);
    }
    const reg = /(\d+)(\w)/g;
    const m = 60;
    const h = m * 60;
    sum = 0;
    const arr = time.match(reg);

    for (var a of arr) {
        if (arr[0]) {
            n = arr[0];
            let value = parseInt(n.slice(0, -1));
            var total = sum + (h * value)

        }
        if (arr[1]) {
            n = arr[1];
            const key = n.slice(-1);
            let value = parseInt(n.slice(0, -1));
            var sum = total + (m * value)
        }
        return sum;

    }


}