   /*처음 시작시 자동 숨김 QR코드*/
   $(function(){
    $("#hide").hide(0.1,"linear");
  });
/*처음 시작시 자동 숨김 왼쪽 메뉴 버튼*/
$(function(){
    $("#hide2").hide(0.1,"linear");
  });
/*포인트 목록 자동 숨김*/
$(function(){
    $("#hide3").hide(0.1,"linear");
  });
/*이벤트 목록 자동 숨김*/
$(function(){
    $("#hide4").hide(0.1,"linear");
  });

/*아래 QR코드 클릭시 QR코드 보임*/
 $(function(){
      $(".ghost-button2").on("click",function(){
          $("#hide").slideToggle(10,"linear");
          });
  });
 /*왼쪽 메뉴 클릭시 메뉴 보임*/
 $(function(){
     $(".ghost-button").on("click",function(){
         $("#hide2").slideToggle(10,"linear");
         });
 });
 /*QR코드 취소시 화면 숨김*/
 $(function(){
     $(".cancel1").on("click",function(){
         $("#hide").slideToggle(10,"linear");

         });
 });
 /*왼쪽 메뉴 취소시 화면 숨김*/
  $(function(){
      $(".cancel2").on("click",function(){
          $("#hide2").slideToggle(10,"linear");

          });
});
 /*포인트 목록 클릭 시 보임*/
 $(function(){
  $(".point_click").on("click",function(){
      $("#hide3").slideToggle(10,"linear");

      });
});
 /**/
$(function(){
  $(".event_click").on("click",function(){
      $("#hide4").slideToggle(10,"linear");

      });
});