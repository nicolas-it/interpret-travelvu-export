import readXlsxFile from "read-excel-file/node";
import writeXlsxFile from 'write-excel-file/node'
import { isEqual, differenceInMinutes, addDays } from "date-fns";
import { exit } from "process";
import * as path from "path";

type Activity = {
    mode?: string;
    duration: number;
    distance: number;
};

if (process.argv.length < 3) {
    console.error("input file missing");
    exit(255);
}

const filePath = process.argv[2];
const outFile = path.join(path.dirname(filePath), path.basename(filePath, path.extname(filePath)) + "-result" + path.extname(filePath));

// duration are strangely formated as TimeStamp starting at the following date
const startDurationDate = new Date("1899-12-30T00:00:00.000Z");
const activityList: {
    date: Date;
    activities: Activity[]
}[] = [];

const minutesToReadableDuration = (minutes: number): string => {
    const hours = Math.floor(minutes / 60);

    return `${hours}:${(minutes - hours * 60).toString().padStart(2, "0")}`;
};

readXlsxFile(filePath).then(async rows => {
    let rowCount = 0;

    for (const row of rows) {
        // skip first row (just headers)
        if (rowCount++ > 0 && row[0]) {
            const activityDate = row[0] as any as Date;
            const mode = row[2]?.toString();
            const duration = differenceInMinutes(row[4] as any, startDurationDate);
            const distance = Number(row[5]?.toString().replace(",", ".") ?? 0);

            const activity: Activity = {
                mode,
                duration,
                distance
            }

            const position = activityList.find(pos => isEqual(pos.date, activityDate));


            if (position) {
                position.activities.push(activity);
            } else {
                activityList.push({
                    date: activityDate,
                    activities: [activity]
                })
            }
        }
    }


    const data: any[] = [[
        {
            value: "Tag",
        },
        {
            value: "Datum",
        },
        {
            value: "Ohne Aufzeichnung"
        },
        {
            value: "Gehen_Dauer"
        },
        {
            value: "Gehen_Distanz"
        },
        {
            value: "Radfahren_Dauer"
        },
        {
            value: "Radfahren_Distanz"
        },
        {
            value: "Summe_Dauer_AT"
        },
        {
            value: "Summe_Distanz_ AT"
        },
        {
            value: "Fahren_Dauer"
        },
        {
            value: "Fahren_Distanz"
        },
        {
            value: "Mitfahren_Dauer"
        },
        {
            value: "Mitfahren_Distanz"
        },
        {
            value: "Summe_Dauer_PT"
        },
        {
            value: "Summe_Distanz_ PT"
        }]
    ];



    let day = 1;
    let startDate = activityList[0].date;
    let activityDay: { date: Date; activities: Activity[]; } | undefined;
    let activityDaysCount = 0;

    do {
        let walkDuration = 0;
        let walkDistance = 0;
        let bicycleDuration = 0;
        let bicycleDistance = 0;
        let sumActiveDuration = 0;
        let sumActiveDistance = 0;
        let driveDuration = 0;
        let driveDistance = 0;
        let passengerDuration = 0;
        let passengerDistance = 0;
        let sumPassiveDuration = 0;
        let sumPassiveDistance = 0;

        const dateCounter = addDays(startDate, day - 1);
        
        activityDay = activityList.find(pos => isEqual(pos.date, dateCounter))

        if (activityDay) {
            activityDaysCount++;

            for (const activity of activityDay.activities) {
                switch (activity.mode?.toLowerCase()) {
                    case "walk":
                        walkDistance += activity.distance;
                        walkDuration += activity.duration;
                        break;
                    case "bicycle":
                    case "electric bicycle":
                        bicycleDistance += activity.distance;
                        bicycleDuration += activity.duration;
                        break;
                    case "car":
                    case "moped":
                    case "electric scooter":
                        driveDistance += activity.distance;
                        driveDuration += activity.duration;
                        break;
                    case "bus":
                    case "car passenger":
                    case "train":
                    case "metro":
                    case "taxi":
                    case "tram":
                    case "airplane":
                    case "ferry/boat":
                        passengerDistance += activity.distance;
                        passengerDuration += activity.duration;
                        break;
                }
            }

            sumActiveDuration = walkDuration + bicycleDuration;
            sumActiveDistance = walkDistance + bicycleDistance;

            sumPassiveDuration = driveDuration + passengerDuration;
            sumPassiveDistance = driveDistance + passengerDistance;
        }

        const row = [
            {
                type: Number,
                value: day,
                align: "right"
            }, {
                type: Date,
                value: dateCounter,
                align: "right",
                format: 'dd.mm.yyyy'
            }, {
                type: String,
                value: activityDay ? '' : 'X',
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(walkDuration),
                align: "right"
            }, {
                type: Number,
                value: walkDistance,
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(bicycleDuration),
                align: "right"
            }, {
                type: Number,
                value: bicycleDistance,
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(sumActiveDuration),
                align: "right"
            }, {
                type: Number,
                value: sumActiveDistance,
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(driveDuration),
                align: "right"
            }, {
                type: Number,
                value: driveDistance,
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(passengerDuration),
                align: "right"
            }, {
                type: Number,
                value: passengerDistance,
                align: "right"
            }, {
                type: String,
                value: minutesToReadableDuration(sumPassiveDuration),
                align: "right"
            }, {
                type: Number,
                value: sumPassiveDistance,
                align: "right"
            }
        ];

        data.push(row);
        day++;

    } while (!isEqual(activityDay?.date ?? 0, activityList[activityList.length - 1].date));

    data.unshift([
        {
            type: String,
            value: "Aktivitätstage"
        }, {
            type: Number,
            value: activityDaysCount
        }
    ], []);


    await writeXlsxFile(data, {
        columns: [{
            width: 12
        }, {
            width: 12
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }, {
            width: 15
        }],
        filePath: outFile
    })
})

