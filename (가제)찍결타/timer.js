function timer(){
    var second = 0, minute = 0, hour = 0;
              setInterval(count, 1000)

              function count() {
                     if (second < 59) {
                            second += 1;
                     } else {
                            second = 0;
                            minute += 1;
                     }

                     if (minute >= 60) {
                            minute = 0;
                            hour += 1;
                     }

                     document.getElementById("timer").innerHTML = hour + ':' + minute + ':' + second;
              }
}