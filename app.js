document.addEventListener('DOMContentLoaded', () => {
    // --- O'zgaruvchilarni e'lon qilish ---
    const steps = document.querySelectorAll('.wizard-step');
    const stepItems = document.querySelectorAll('.step-item');
    const nextBtns = {
        to2: document.getElementById('toStep2'),
        to3: document.getElementById('toStep3'),
        to4: document.getElementById('toStep4')
    };
    const backBtns = {
        to1: document.getElementById('backToStep1'),
        to2: document.getElementById('backToStep2'),
        to3: document.getElementById('backToStep3')
    };

    // Ma'lumotlar ombori
    let userData = {
        personal: {},
        profileImage: null,
        projects: [],
        selectedTemplate: 'classic'
    };

    // --- 1. Bosqichma-bosqich o'tish logikasi ---
    function goToStep(stepNumber) {
        steps.forEach(step => step.classList.remove('active'));
        stepItems.forEach(item => {
            item.classList.remove('active');
            if (parseInt(item.dataset.step) <= stepNumber) {
                item.classList.add('active');
            }
        });
        document.getElementById(`step${stepNumber}`).classList.add('active');
        window.scrollTo(0, 0);
    }

    nextBtns.to2.addEventListener('click', () => {
        const form = document.getElementById('personalForm');
        if (form.checkValidity()) {
            userData.personal = {
                name: `${document.getElementById('firstName').value} ${document.getElementById('lastName').value}`,
                phone: document.getElementById('phone').value,
                profession: document.getElementById('profession').value,
                bio: document.getElementById('bio').value
            };
            goToStep(2);
        } else {
            form.reportValidity();
        }
    });

    nextBtns.to3.addEventListener('click', () => goToStep(3));
    nextBtns.to4.addEventListener('click', () => {
        renderPreview();
        goToStep(4);
    });

    backBtns.to1.addEventListener('click', () => goToStep(1));
    backBtns.to2.addEventListener('click', () => goToStep(2));
    backBtns.to3.addEventListener('click', () => goToStep(3));

    // --- 2. Profil rasm yuklash ---
    const profileUploadArea = document.getElementById('profileUploadArea');
    const profileInput = document.getElementById('profilePhotoInput');
    const profilePreview = document.getElementById('profilePreview');

    profileUploadArea.addEventListener('click', () => profileInput.click());

    profileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (event) => {
                userData.profileImage = event.target.result;
                profilePreview.innerHTML = `<img src="${userData.profileImage}" style="width:100%; height:100%; border-radius:50%; object-fit:cover;">`;
            };
            reader.readAsDataURL(file);
        }
    });

    // --- 3. Loyihalar qo'shish (Max 10) ---
    const projectsContainer = document.getElementById('projectsContainer');
    const addProjectBtn = document.getElementById('addProjectBtn');
    const projectCountLabel = document.getElementById('projectCount');

    addProjectBtn.addEventListener('click', () => {
        if (userData.projects.length >= 10) {
            showToast("Maksimal 10 ta loyiha qo'shish mumkin!");
            return;
        }

        const projectCard = document.createElement('div');
        projectCard.className = 'project-upload-card'; // CSSda buni shakllantiring
        projectCard.innerHTML = `
            <div class="project-input-group">
                <input type="file" class="proj-img-input" accept="image/*">
                <input type="text" class="proj-title-input" placeholder="Loyiha nomi">
                <button class="remove-proj"><i class="fas fa-trash"></i></button>
            </div>
        `;
        projectsContainer.appendChild(projectCard);
        updateProjectCount();

        // Rasm yuklash va o'chirish hodisalari
        const fileInput = projectCard.querySelector('.proj-img-input');
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = (ev) => {
                    const index = Array.from(projectsContainer.children).indexOf(projectCard);
                    userData.projects[index] = { 
                        img: ev.target.result, 
                        title: projectCard.querySelector('.proj-title-input').value 
                    };
                };
                reader.readAsDataURL(file);
            }
        });

        projectCard.querySelector('.remove-proj').addEventListener('click', () => {
            projectCard.remove();
            updateProjectCount();
        });
    });

    function updateProjectCount() {
        const count = projectsContainer.children.length;
        projectCountLabel.innerText = `(${count}/10)`;
    }

    // --- 4. Shablon tanlash ---
    const templateCards = document.querySelectorAll('.template-card');
    templateCards.forEach(card => {
        card.addEventListener('click', () => {
            templateCards.forEach(c => c.classList.remove('selected'));
            card.classList.add('selected');
            userData.selectedTemplate = card.dataset.template;
            nextBtns.to4.disabled = false;
        });
    });

    // --- 5. Preview ko'rsatish ---
    function renderPreview() {
        const preview = document.getElementById('portfolioPreview');
        preview.className = `portfolio-preview template-${userData.selectedTemplate}`;
        
        preview.innerHTML = `
            <div class="preview-header">
                <h2>${userData.personal.name}</h2>
                <p>${userData.personal.profession}</p>
            </div>
            <div class="preview-body">
                ${userData.projects.map(p => `
                    <div class="preview-item">
                        <img src="${p.img}" alt="">
                        <span>${p.title}</span>
                    </div>
                `).join('')}
            </div>
        `;
    }

    // --- 6. PowerPoint yaratish (PptxGenJS) ---
    document.getElementById('downloadPptx').addEventListener('click', () => {
        const loader = document.getElementById('loadingOverlay');
        loader.classList.add('active');

        let pptx = new PptxGenJS();

        // 1-Slayd: Kirish
        let slide1 = pptx.addSlide();
        slide1.addText(userData.personal.name, { x: 1, y: 1, fontSize: 32, bold: true, color: '363636' });
        slide1.addText(userData.personal.profession, { x: 1, y: 1.5, fontSize: 20, color: '666666' });
        if (userData.profileImage) {
            slide1.addImage({ data: userData.profileImage, x: 6, y: 1, w: 3, h: 3 });
        }

        // Loyihalar slaydlari
        userData.projects.forEach(proj => {
            let slide = pptx.addSlide();
            slide.addText(proj.title, { x: 0.5, y: 0.5, fontSize: 24, bold: true });
            if (proj.img) {
                slide.addImage({ data: proj.img, x: 0.5, y: 1.2, w: 9, h: 4.5 });
            }
        });

        pptx.writeFile({ fileName: `${userData.personal.name}_Portfolio.pptx` })
            .then(() => {
                loader.classList.remove('active');
                showToast("Portfolio yuklab olindi!");
            });
    });

    // Toast funksiyasi
    function showToast(msg) {
        const toast = document.getElementById('toast');
        document.getElementById('toastMessage').innerText = msg;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 3000);
    }
});