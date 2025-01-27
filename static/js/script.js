// Grafik oluşturma fonksiyonu
function generateChart() {
  const xCoords = document
    .getElementById("x-coordinates")
    ?.value.split(",")
    .map((item) => item.trim());
  const cevaplananCagri = document
    .getElementById("cevaplanan-cagri")
    ?.value.split(",")
    .map((item) => (item.trim() === "" ? 0 : parseFloat(item.trim())));
  const gelenCagri = document
    .getElementById("gelen-cagri")
    ?.value.split(",")
    .map((item) => (item.trim() === "" ? 0 : parseFloat(item.trim())));

  // Hata kontrolü
  if (!xCoords || !cevaplananCagri || !gelenCagri) {
    alert("Bir veya daha fazla gerekli giriş alanı bulunamadı. Lütfen sayfayı kontrol edin!");
    return;
  }

  if (
    xCoords.length === 0 ||
    cevaplananCagri.length === 0 ||
    gelenCagri.length === 0 ||
    xCoords.length !== cevaplananCagri.length ||
    cevaplananCagri.length !== gelenCagri.length
  ) {
    alert("Lütfen tüm alanları doldurun ve verilerinizi doğru formatta girin!");
    return;
  }

  const oranlar = cevaplananCagri.map(
    (cevap, index) => (cevap / gelenCagri[index]) * 100
  );

  const ctx = document.getElementById("myChart").getContext("2d");
  if (window.myChart instanceof Chart) {
    window.myChart.destroy();
  }

  window.myChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: xCoords,
      datasets: [
        {
          label: "Toplam Gelen Çağrı",
          data: gelenCagri,
          backgroundColor: "rgba(153, 102, 255, 0.5)",
          borderColor: "rgba(153, 102, 255, 1)",
          borderWidth: 1,
          yAxisID: "y",
        },
        {
          label: "Toplam Cevaplanan Çağrı",
          data: cevaplananCagri,
          backgroundColor: "rgba(75, 192, 192, 0.5)",
          borderColor: "rgba(75, 192, 192, 1)",
          borderWidth: 1,
          yAxisID: "y",
        },
        {
          type: "line",
          label: "Cevaplanan / Gelen Çağrı Oranı (%)",
          data: oranlar,
          borderColor: "rgba(255, 159, 64, 1)",
          backgroundColor: "rgba(255, 159, 64, 0)",
          fill: false,
          tension: 0,
          borderWidth: 2,
          pointStyle: "circle",
          pointRadius: 5,
          pointBackgroundColor: "white",
          yAxisID: "y1",
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        title: {
          display: true,
          text: "Çağrı Raporu",
        },
        tooltip: {
          mode: "index",
          intersect: false,
        },
        legend: {
          display: true,
        },
      },
      scales: {
        x: {
          stacked: true,
        },
        y: {
          stacked: true,
          beginAtZero: true,
          position: "left",
        },
        y1: {
          beginAtZero: true,
          position: "right",
          title: {
            display: true,
            text: "Cevaplanan / Gelen Çağrı Oranı (%)",
          },
          grid: {
            drawOnChartArea: false,
          },
        },
      },
    },
  });
}

// Excel dosyasını dışa aktarma fonksiyonu
function exportToExcel() {
  const xCoords = document
    .getElementById("x-coordinates")
    .value.split(",")
    .map((item) => item.trim());
  const cevaplananCagri = document
    .getElementById("cevaplanan-cagri")
    .value.split(",")
    .map((item) => parseFloat(item.trim()));
  const gelenCagri = document
    .getElementById("gelen-cagri")
    .value.split(",")
    .map((item) => parseFloat(item.trim()));

  if (xCoords.length === 0 || cevaplananCagri.length === 0 || gelenCagri.length === 0) {
    alert("Lütfen tüm alanları doldurun ve verilerinizi doğru girin.");
    return;
  }

  fetch('/generate_excel', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      x_coords: xCoords,
      cevaplanan: cevaplananCagri,
      gelen: gelenCagri,
    }),
  })
    .then((response) => response.blob())
    .then((blob) => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'veriler.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch((error) => {
      console.error('Excel oluşturulurken bir hata oluştu:', error);
    });
}

// Excel dosyasını yükle
function uploadExcel() {
  const fileInput = document.getElementById("excelFile");
  const file = fileInput.files[0];
  const fileNameDisplay = document.getElementById("fileNameDisplay");

  if (!file) {
    alert("Lütfen bir Excel dosyası seçin.");
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  fetch("http://127.0.0.1:5000/upload_excel", {  // Yerel geliştirme için URL
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (!response.ok) {
        throw new Error("Sunucudan hata alındı.");
      }
      return response.json();
    })
    .then((data) => {
      if (!data.success) {
        alert(data.message || "Dosya yüklenirken bir hata oluştu.");
        fileNameDisplay.textContent = "Yükleme Başarısız!";
        return;
      }

      const { x_coords, toplam_gelen, toplam_cevaplanan } = data.data;

      // Verileri giriş alanlarına aktar
      if (x_coords && x_coords.length > 0) {
        document.getElementById("x-coordinates").value = x_coords.join(",");
      }
      if (toplam_gelen && toplam_gelen.length > 0) {
        document.getElementById("gelen-cagri").value = toplam_gelen.join(",");
      }
      if (toplam_cevaplanan && toplam_cevaplanan.length > 0) {
        document.getElementById("cevaplanan-cagri").value = toplam_cevaplanan.join(",");
      }

      fileNameDisplay.textContent = `Yüklenen Dosya: ${file.name}`;
      alert("Excel dosyasındaki veriler başarıyla yüklendi!");
    })
    .catch((error) => {
      console.error("Dosya yüklenirken bir hata oluştu:", error);
      fileNameDisplay.textContent = "Yükleme Başarısız!";
      alert("Dosya yüklenirken bir hata oluştu. Lütfen tekrar deneyin.");
    });
}

// Seçilen dosyanın adını göster
function showFileName() {
  const fileInput = document.getElementById("excelFile");
  const fileNameDisplay = document.getElementById("fileNameDisplay");

  if (fileInput.files && fileInput.files[0]) {
    fileNameDisplay.textContent = `Seçilen Dosya: ${fileInput.files[0].name}`;
  } else {
    fileNameDisplay.textContent = "Henüz bir dosya seçilmedi.";
  }
}

// Makro kodu kopyalama
function showMacroCode() {
  const macroCodeElement = document.getElementById("macroCode");
  const macroCode = macroCodeElement.textContent || macroCodeElement.innerText;

  navigator.clipboard
    .writeText(macroCode)
    .then(() => {
      alert("Makro kodu kopyalandı!");
    })
    .catch((error) => {
      console.error("Makro kodu kopyalanırken bir hata oluştu:", error);
      alert("Makro kodu kopyalanamadı.");
    });
}

// Pivot grafik rehberi
function addPivotChart() {
  alert(
    "Excelde pivot grafik elde etmek için sırası ile bu adımları takip edicez. \n" +
      "1- Excel'de Alt tuşuna ve F11 tuşuna aynı anda bas \n" +
      "2- Açılan ekranda sol üst tarafta 'Insert' başlığı altındaki 'Module' sekmesine gir \n" +
      "3- Web sayfasında bulunan 'Makro Kodunu Göster' butonuna bas ve gelen kodu kopyala (Excel VBA Makro Kodu: yazan başlık hariç) \n" +
      "4- 'Module' sayfasına kopyaladığın kodu yapıştır \n" +
      "5- Tekrardan Alt ve F11 tuşlarına aynı anda basıp Excel arayüzüne geri dön \n" +
      "6- Alt ve F8 tuşlarına aynı anda basıp karşına gelecek olan 'AddOrUpdatePivotChart' seçeneğini seçip 'Çalıştır'a tıkla. Şu anda pivot tablon oluşturulmuş olmalı. \n" +
      "7- Eğer güncelleme yapmak istersen 'Veriler' sayfasından değiştirmek istediğin verileri değiştir. Daha sonradan Alt ve F8'e basıp bu sefer 'RefreshPivotTable' seçeneğini seçip 'Çalıştır'a bas. Bu, pivot tablonu değiştirmeye yarayacak. \n" +
      "Afiyet olsun! 😊"
  );
}
