// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

// CLASS .CLEAR-ON-BLUR XÓA DỮ LIỆU Ô INPUT MỖI KHI THAO TÁC XONG
document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll(".clear-on-blur").forEach(function (el) {
        el.addEventListener("blur", function () {
            this.value = "";
        });
    });
});

// JS CỦA PROTECTDUTY -> INDEX
document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll('.toast').forEach(toastEl => {
        const toast = new bootstrap.Toast(toastEl, {
            delay: 2000, // 2 giây
            autohide: true
        });
        toast.show();
    });
    // JS để load modal Edit qua AJAX
    document.querySelectorAll('.edit-btn').forEach(button => {
        button.addEventListener('click', function () {
            const id = this.getAttribute('data-id');
            const url = '/ProtectDuty/Edit/' + id;

            fetch(url)
                .then(response => response.text())
                .then(html => {
                    // Insert HTML của partial view vào placeholder
                    document.getElementById('editModalPlaceholder').innerHTML = html;
                    // Show modal
                    const editModal = new bootstrap.Modal(document.getElementById('editModal'));
                    editModal.show();
                })
                .catch(error => console.error('Lỗi khi load modal Edit:', error));
        });
    });
});

// NGĂN XÓA GIÁ TRỊ Ô FORMDATE VÀ TO DATE TRONG VIEW INOUT
document.addEventListener("DOMContentLoaded", function () {
    document.getElementById('toDate').addEventListener('keydown', function (e) {
        if (e.key === 'Backspace' || e.key === 'Delete') {
            if (e.target.value.length <= 1) {
                e.preventDefault(); // Ngăn hành động xóa
            }
        }
    });

    document.getElementById('toDate').addEventListener('input', function (e) {
        if (!e.target.value) {
            e.target.value = '@toDateTime'; // Khôi phục giá trị mặc định
        }
    });
});


// kích hoạt đóng modal export khi đã xuất excel thành công
document.addEventListener("DOMContentLoaded", function () {
    const form = document.getElementById("closeForm");
    form.addEventListener("submit", function () {
        // Lấy nút Hủy
        const closeBtn = document.getElementById("closeButton");
        if (closeBtn) {
            closeBtn.click(); // kích hoạt nút để đóng modal
        }
    });
});