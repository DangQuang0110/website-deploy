document.getElementById("exportWord").addEventListener("click", async function () {
    try {
        // Fetch the external CSS file
        const cssResponse = await fetch('static/style.css');
        const cssText = await cssResponse.text();
        
        // Wrap the CSS in style tags
        const styles = `<style>${cssText}</style>`;

        const contentElement = document.getElementById("exportContent");
        let contentHTML = contentElement.innerHTML;

        const images = contentElement.querySelectorAll('img');

        for (let img of images) {
            try {
                const imgSrc = img.getAttribute('src');

                if (imgSrc && !imgSrc.startsWith('data:')) {
                    // Fetch the image
                    const response = await fetch(imgSrc);
                    const blob = await response.blob();

                    // Convert to base64
                    const base64 = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result);
                        reader.readAsDataURL(blob);
                    });

                    // Replace image source with base64 data
                    contentHTML = contentHTML.replace(imgSrc, base64);
                }
            } catch (imgError) {
                console.error("Error processing image:", imgError);
            }
        }
        
        const header = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" 
                  xmlns:w="urn:schemas-microsoft-com:office:word" 
                  xmlns="http://www.w3.org/TR/REC-html40">
            <head>
                <meta charset="utf-8">
                <title>CV - Phạm Đăng Quang</title>
                ${styles}
            </head>
            <body>`;
        const footer = `</body></html>`;
        const sourceHTML = header + contentHTML + footer;

        const blob = new Blob(['\ufeff', sourceHTML], { 
            type: 'application/msword'
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "CV_PhamDangQuang.doc";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Clean up the URL object after download starts
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);

    } catch (error) {
        console.error("Error generating Word document:", error);
        alert("There was an error generating the Word document. Please try again.");
    }
});