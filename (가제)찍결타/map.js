/* 지도 생성 -> 현재 위치를 기준으로 초기화한다.
   핀 위치 -> 현재 근처 기기들의 위치를 표시한다.
   핀을 클릭하면 그 기기에 대한 인포 창(기기상태, 배터리 등)을 보여주고
   마우스를 이동하면(안드로이드라면 다른곳을 터치하면..?) 인포 창이 사라진다. */
var map;

function map_Initial() {
       var container = document.getElementById('map');
       var options = {
              // 맵 중앙 위치를 나타냅니다
              center: new kakao.maps.LatLng(35.116053, 128.967356),
              // 맵 확대 레벨입니다
              level: 2
       };

       map = new kakao.maps.Map(container, options);

       // 기기 위치를 나타냅니다
       var positions = [
    	   	  {
    	   			content: '<div id="info_Box"><input type="text" value="현재 위치"></div>',
    	   			latlng: new kakao.maps.LatLng(35.115518, 128.967536)
    	   	  },
              {
                     content: '<div id="info_Box"><textarea>배터리: 48%\n거리: 23m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116239, 128.967560)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 63%\n거리: 42m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116365, 128.967952)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 4% (충전이 필요합니다)\n거리: 43m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.115857, 128.966621)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 25%\n거리: 8m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.115651, 128.967771)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 27%\n거리: 9m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.115525, 128.967908)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 86%\n거리: 50m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116502, 128.967771)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 94%\n거리: 38m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116340, 128.967727)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 63%\n거리: 20m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116181, 128.967441)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 98%\n거리: 25m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.115994, 128.966881)
              },
              {
                     content: '<div id="info_Box"><textarea>배터리: 74%\n거리: 23m</textarea></div>',
                     latlng: new kakao.maps.LatLng(35.116044, 128.967013)
              }
       ];


       var myLoc = "http://t1.daumcdn.net/localimg/localimages/07/mapapidoc/markerStar.png";
       
       // 마커 이미지의 이미지 주소입니다
       var imageSrc = "http://t1.daumcdn.net/localimg/localimages/07/mapapidoc/marker_red.png";

       
       
       for (var i = 0; i < positions.length; i++) {

              // 마커 이미지의 이미지 크기 입니다
              var imageSize = new kakao.maps.Size(25, 35);

              // 마커 이미지를 생성합니다
              var markerImage = new kakao.maps.MarkerImage(imageSrc, imageSize);
              if(i == 0){
            	  markerImage = new kakao.maps.MarkerImage(myLoc, imageSize);
              }

              // 마커를 생성합니다
              var marker = new kakao.maps.Marker({
                     map: map, // 마커를 표시할 지도
                     position: positions[i].latlng, // 마커를 표시할 위치
                     image: markerImage, // 마커 이미지
                     clickable: true
              });

              // 클릭하면 띄울 인포윈도우입니다
              var infowindow = new kakao.maps.InfoWindow({
                     content: positions[i].content, // 인포윈도우에 표시할 내용
                     removable: true
              });

              // 클릭 시, 인포윈도우 창 생성
              kakao.maps.event.addListener(marker, 'click', makeOverListener(map, marker, infowindow));
              // 마우스를 옮기기만 해도 인포윈도우 창 제거
              //kakao.maps.event.addListener(marker, 'mouseout', makeOutListener(infowindow));
       }


       // 인포윈도우를 표시하는 클로저를 만드는 함수입니다
       function makeOverListener(map, marker, infowindow) {
              return function () {
                     infowindow.open(map, marker);
              };
       }

       // 인포윈도우를 닫는 클로저를 만드는 함수입니다
       function makeOutListener(infowindow) {
              return function () {
                     infowindow.close();
              };
       }
}

function moveTo() {
       // 이동할 위도 경도 위치를 생성합니다
       var moveLatLon = new kakao.maps.LatLng(35.116053, 128.967356);

       // 지도 중심을 부드럽게 이동시킵니다
       // 만약 이동할 거리가 지도 화면보다 크면 부드러운 효과 없이 이동합니다
       map.panTo(moveLatLon);
}