// Grafik oluşturma fonksiyonu
function generateChart() {
  const xCoords = document
    .getElementById("x-coordinates")
    ?.value.split(",")
    .map((item) => item.trim());
  const cevaplananCagri = document
    .getElementById("cevaplanan-cagri")
    ?.value.split(",")
    .map((item) => parseFloat(item.trim()));
  const gelenCagri = document
    .getElementById("gelen-cagri")
    ?.value.split(",")
    .map((item) => parseFloat(item.trim()));

  // Hata kontrolü
  if (!xCoords || !cevaplananCagri || !gelenCagri) {
    alert("Bir veya daha fazla alan eksik!");
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

// Dosya ismini gösterme fonksiyonu
function showFileName(event) {
  // event.target ile dosya ismini almak
  const fileInput = event.target;
  const fileName = fileInput.files[0]?.name;

  if (fileName) {
    const fileNameElement = document.getElementById("fileNameDisplay");
    if (fileNameElement) {
      fileNameElement.textContent = fileName;
    }
  } else {
    const fileNameElement = document.getElementById("fileNameDisplay");
    if (fileNameElement) {
      fileNameElement.textContent = "Henüz bir dosya seçilmedi.";
    }
  }
}




// Excel'e veri aktarımı fonksiyonu
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

  // Backend'e JSON verilerini gönder
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
      // Excel dosyasını indirme
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

// Excel dosyasını yükleme fonksiyonu
function uploadExcel() {
  const fileInput = document.getElementById("excelFile");
  const file = fileInput.files[0];

  if (!file) {
    alert("Lütfen bir Excel dosyası seçin.");
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  fetch("https://mypivot-4q9n.onrender.com/", {
    method: "POST",
    body: formData,
  })
    .then((response) => response.json())
    .then((data) => {
      if (!data.success) {
        alert(data.message || "Dosya yüklenirken bir hata oluştu.");
        return;
      }

      const { x_coords, toplam_gelen, toplam_cevaplanan } = data.data;

      // X Koordinatlarını doldur
      if (x_coords && x_coords.length > 0) {
        document.getElementById("x-coordinates").value = x_coords.join(",");
      }

      // Diğer alanları doldur
      if (toplam_gelen && toplam_gelen.length > 0) {
        document.getElementById("gelen-cagri").value = toplam_gelen.join(",");
      }

      if (toplam_cevaplanan && toplam_cevaplanan.length > 0) {
        document.getElementById("cevaplanan-cagri").value = toplam_cevaplanan.join(",");
      }

      alert("Excel dosyasındaki veriler başarıyla yüklendi!");
    })
    .catch((error) => {
      console.error("Dosya yüklenirken bir hata oluştu:", error);
      alert("Dosya yüklenirken bir hata oluştu. Lütfen tekrar deneyin.");
    });
}

// Makro kodunu kopyalama fonksiyonu
function showMacroCode() {
  const macroCodeElement = document.getElementById("macroCode");
  const macroCode = macroCodeElement.textContent || macroCodeElement.innerText;

  // Makro kodunu kopyala
  navigator.clipboard.writeText(macroCode).then(() => {
    alert("Makro kodu kopyalandı!");
  }).catch((error) => {
    console.error("Makro kodu kopyalanırken bir hata oluştu:", error);
    alert("Makro kodu kopyalanamadı.");
  });
}

// Pivot grafik rehberi fonksiyonu
function addPivotChart() {
  alert("Excelde pivot grafik elde etmek için sırası ile bu adımları takip edicez. \n 1-Excelde Alt tuşuna ve F11 tuşuna aynı anda bas \n 2-Açılan ekranda sol üst tarafta Insert başlığı altındaki modul sekmesine gir \n 3-Web sayfasında bulunan makro kodunu göster butonuna basınca gelen kodu kopyala (Excel VBA Makro Kodu: yazan başlık hariç) \n 4-Module sayfasına kopyaladığın kodu yapıştır \n 5-Tekrardan Alt ve F11 tuşlarına aynı anda basıp excel arayüzüne geri dön \n 6-Alt ve F8 tuşlarına aynı anda basıp karşına gelecek olan AddOrUpdatePivotChart seçeneğini seçip çalıştıra tıkla. Şuanda pivot tablon oluşturulmuş olmalı \n 7- Eğer güncelleme yapmak istersen veriler sayfasından değiştirmek istediğin verileri değiştir, daha sonradan Alt ve F8'e basıp bu seferde RefreshPivotTable seçeneğini seçip çalıştıra bas, bu senin pivot tablonu değiştirmeye yarayacak \n Afiyet olsun! :+D");
}
