$(document).ready(function () {
    // 加入 CSS loading 
    const loadingHTML = `
        <div id="custom-loading-overlay"
            style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;
                    background-color:rgba(204, 204, 204, 0.8);z-index:9999;">
            <div id="loading-inner">
            <div class="custom-loading-spinner"
                style="border:18px solid #f3f3f3;border-top:18px solid #a2d8fc;border-radius:50%;
                        width:100px;height:100px;animation:spin 1s linear infinite;"></div>
            <div style="margin-top:20px;color:#535151;font-size:20px;">處理中</div>
            </div>
        </div>
    `;
    $("body").append(loadingHTML);
  
    const style = `
        <style>
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        </style>
    `;
    $("head").append(style);

    // API URL
    var API_URL = "https://script.google.com/macros/s/AKfyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/exec";
    $("#ITEM310").click(function () {
        var testImgurl = $("textarea#ITEM365").val().trim();
        var sampleData = {
            ticketNo: $("#ITEM5").val(),
            aplyDate: $("#ITEM24").val(),
            abnItem: $("#ITEM88").val(),
            materialName: $("#ITEM56").val(),
            descript: $("#ITEM186").val(),
            imageUrl: testImgurl,
            vendorName: $("#ITEM130").val(),
            vendorEmail: $("#ITEM327").val(),
            notificationEmail: $("#ITEM319").val(),
            notifyQaEmail: $("#ITEM337").val(),
        };

        var fieldLabels = {
            ticketNo: "單號",
            aplyDate: "申請日期",
            abnItem: "異常項目",
            materialName: "物料名稱",
            descript: "描述",
            vendorName: "供應商名稱",
            vendorEmail: "供應商 Email",
            notificationEmail: "責任人員Email",
            notifyQaEmail: "QA執行人員Email",
        };

        if (!validateFields(sampleData, fieldLabels)) {
            return; // 有欄位沒填，終止送出
        }
        showLoading();

        //post
        fetch(API_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(sampleData),
        })
            .then((res) => res.json())
            .then((data) => {
                hideLoading();
                if (confirm("問卷已建立，是否開啟問卷？\n\n" + data.formUrl)) {
                    window.open(data.formUrl, "_blank");
                }
                console.log(data.formUrl);

                //更新-追蹤狀態 觸發儲存表單
                $("#ITEM338").val("待廠商回覆").trigger("change");
            })
            .catch((err) => {
                hideLoading();
                console.log(err.message);
                alert("建立失敗，請重新產生問券：" + err.message);
            });
    });

    $("#ITEM311").click(function () {
        var ticketNo = $("#ITEM5").val();
        if (ticketNo == "") {
            return;
        }

        showLoading();
        fetch(`${API_URL}?ticketNo=${ticketNo}`)
            .then((res) => res.json())
            .then((data) => {
                hideLoading();
                if (data.status === "success") {
                    const d = data.data;
                    alert("查詢成功：" + d.ticketNo);

                    console.log(d);
                    $("#ITEM214").val(d.reply1);
                    $("#ITEM217").val(d.reply2);
                    $("#ITEM223").val(d.reply3);
                    $("#ITEM229").val(d.reply4);
                    $("#ITEM338").val("完成追蹤");
                } else {
                    alert(ticketNo + "查詢失敗：無資料");
                }
            })
            .catch((err) => {
                alert("查詢錯誤：" + err.message);
            });
    });

    function validateFields(data, fieldLabels) {
        // 定義非必填欄位
        const optionalFields = ["vendorName"];

        for (const [key, label] of Object.entries(fieldLabels)) {
            if (optionalFields.includes(key)) continue; // 跳過非必填欄位

            const value = data[key];
            if (value == null || value == undefined || value == "") {
                alert(`請填寫「${label}」`);
                return false;
            }
        }
        return true;
    }

    function showLoading() {
        const $overlay = $("#custom-loading-overlay");
        const $inner = $("#loading-inner");

        // 顯示 overlay 並先設定為透明
        $overlay.css({
            display: "block",
            opacity: 0,
        });

        const centerTop = $("input#ITEM310").offset().top;

        // 將 loading-inner 置中（絕對位置）
        $inner.css({
            position: "absolute",
            top: centerTop + "px",
            left: "50%",
            transform: "translateX(-50%)",
            display: "flex",
            "flex-direction": "column",
            "align-items": "center",
        });

        // 延遲淡入
        setTimeout(() => {
            $overlay.css("opacity", 1);
        }, 10);
    }

    function hideLoading() {
        const $overlay = $("#custom-loading-overlay");
        $overlay.css("opacity", "0"); // 淡出
        setTimeout(() => {
            $overlay.css("display", "none");
        }, 300); // 等待過渡時間後再關掉
    }
});
