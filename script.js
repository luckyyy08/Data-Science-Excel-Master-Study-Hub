// AUDI Interactive Application Logic

document.addEventListener('DOMContentLoaded', () => {
  initNavbar();
  initCarFilters();
  initRangeCalculator();
  initSpecModal();
  initSmoothScroll();
});

// Header scroll effect
function initNavbar() {
  const header = document.querySelector('.header');
  window.addEventListener('scroll', () => {
    if (window.scrollY > 40) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
  });
}

// Vehicle Category Filtering
function initCarFilters() {
  const filterBtns = document.querySelectorAll('.tab-btn');
  const carCards = document.querySelectorAll('.car-card');

  filterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      // Remove active class from all buttons
      filterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const filterValue = btn.getAttribute('data-filter');

      carCards.forEach(card => {
        const category = card.getAttribute('data-category');
        if (filterValue === 'all' || category.includes(filterValue)) {
          card.style.display = 'flex';
          setTimeout(() => {
            card.style.opacity = '1';
            card.style.transform = 'translateY(0)';
          }, 50);
        } else {
          card.style.opacity = '0';
          card.style.transform = 'translateY(15px)';
          setTimeout(() => {
            card.style.display = 'none';
          }, 300);
        }
      });
    });
  });
}

// Dynamic EV Range Calculator
function initRangeCalculator() {
  const speedInput = document.getElementById('speed-slider');
  const tempInput = document.getElementById('temp-slider');
  
  const speedValDisplay = document.getElementById('speed-val');
  const tempValDisplay = document.getElementById('temp-val');
  const resultDisplay = document.getElementById('calc-range-result');
  
  const modeBtns = document.querySelectorAll('.mode-btn');
  let currentModeMultiplier = 1.0; // Efficiency factor

  if (!speedInput || !tempInput || !resultDisplay) return;

  function updateRange() {
    const speed = parseInt(speedInput.value);
    const temp = parseInt(tempInput.value);

    speedValDisplay.textContent = `${speed} km/h`;
    tempValDisplay.textContent = `${temp}°C`;

    // Formula simulation: Base battery 93 kWh, optimal temp 20C, optimal speed 70km/h
    let baseRange = 482; // km at 70 km/h and 20C
    
    // Speed penalty above 70km/h
    const speedFactor = Math.max(0.5, 1 - (Math.abs(speed - 60) * 0.007));
    
    // Temp penalty outside 15-25C range
    const tempDelta = Math.abs(temp - 22);
    const tempFactor = Math.max(0.75, 1 - (tempDelta * 0.008));
    
    const calculatedRange = Math.round(baseRange * speedFactor * tempFactor * currentModeMultiplier);
    
    // Animate number count
    resultDisplay.textContent = calculatedRange;
  }

  speedInput.addEventListener('input', updateRange);
  tempInput.addEventListener('input', updateRange);

  modeBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      modeBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const mode = btn.getAttribute('data-mode');
      if (mode === 'efficiency') currentModeMultiplier = 1.12;
      else if (mode === 'dynamic') currentModeMultiplier = 0.88;
      else currentModeMultiplier = 1.0; // auto/comfort

      updateRange();
    });
  });

  updateRange();
}

// Interactive Spec Modal Window
function initSpecModal() {
  const modalOverlay = document.getElementById('spec-modal');
  const modalClose = document.getElementById('modal-close');
  const detailBtns = document.querySelectorAll('.view-specs-btn');

  const modalTitle = document.getElementById('modal-car-title');
  const modalSubtitle = document.getElementById('modal-car-subtitle');
  const modalSpecList = document.getElementById('modal-specs-container');

  if (!modalOverlay || !modalClose) return;

  // Car specification dataset
  const carSpecsData = {
    'etron-gt': {
      title: 'Audi e-tron GT quattro',
      subtitle: 'Pure Electric Performance Grand Tourer',
      specs: [
        { label: 'Max Power', value: '350 kW (476 PS) / 530 PS Boost' },
        { label: '0-100 km/h (0-60 mph)', value: '4.1 sec (3.1s RS)' },
        { label: 'Battery Capacity', value: '93.4 kWh Gross' },
        { label: 'DC Fast Charging', value: '270 kW (5-80% in 22.5 min)' },
        { label: 'Top Speed', value: '245 km/h' },
        { label: 'Drivetrain', value: 'Electric quattro AWD' }
      ]
    },
    'rsq8': {
      title: 'Audi RS Q8 Performance',
      subtitle: 'Flagship High-Performance SUV',
      specs: [
        { label: 'Engine', value: '4.0L Twin-Turbo V8 Engine' },
        { label: 'Horsepower', value: '600 HP (591 bhp) / 800 Nm' },
        { label: '0-100 km/h (0-60 mph)', value: '3.8 seconds' },
        { label: 'Top Speed', value: '305 km/h (190 mph)' },
        { label: 'Suspension', value: 'Adaptive Air Suspension Sport' },
        { label: 'Exhaust System', value: 'RS Sport Dual Oval Pipes' }
      ]
    },
    'q3-sportback': {
      title: 'Audi Q3 Sportback',
      subtitle: 'Compact Luxury Crossover Coupe',
      specs: [
        { label: 'Engine', value: '2.0L TFSI Turbocharged 4-Cyl' },
        { label: 'Horsepower', value: '228 HP @ 5000 RPM' },
        { label: 'Transmission', value: '7-Speed S tronic Dual-Clutch' },
        { label: 'Fuel Economy', value: '7.8 L/100 km Combined' },
        { label: 'Cockpit', value: 'Audi Virtual Cockpit Plus 12.3"' },
        { label: 'Sound System', value: 'BANG & OLUFSEN 3D Premium' }
      ]
    },
    'a8-sedan': {
      title: 'Audi A8 L Luxury Sedan',
      subtitle: 'The Pinnacle of Executive Comfort & Tech',
      specs: [
        { label: 'Engine Options', value: '3.0L V6 / 4.0L V8 MHEV' },
        { label: 'Interior Comfort', value: 'Relaxation Seat with Foot Massage' },
        { label: 'Lighting Tech', value: 'Digital Matrix LED Headlights' },
        { label: 'Autonomous Tech', value: 'Audi AI Remote Park Assist' },
        { label: 'Horsepower', value: '453 HP' },
        { label: 'Drive', value: 'quattro with Sport Differential' }
      ]
    }
  };

  detailBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const carKey = btn.getAttribute('data-car-key');
      const data = carSpecsData[carKey];

      if (data) {
        modalTitle.textContent = data.title;
        modalSubtitle.textContent = data.subtitle;
        
        modalSpecList.innerHTML = data.specs.map(s => `
          <div class="modal-spec-item">
            <div class="label">${s.label}</div>
            <div class="value">${s.value}</div>
          </div>
        `).join('');

        modalOverlay.classList.add('open');
      }
    });
  });

  modalClose.addEventListener('click', () => {
    modalOverlay.classList.remove('open');
  });

  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) {
      modalOverlay.classList.remove('open');
    }
  });
}

// Smooth Scrolling & Interactive Notification
function initSmoothScroll() {
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
      const targetId = this.getAttribute('href');
      if (targetId === '#') return;
      
      const targetElement = document.querySelector(targetId);
      if (targetElement) {
        e.preventDefault();
        targetElement.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }
    });
  });
}

// Toast notification helper for CTA buttons
function showToast(message) {
  const toast = document.createElement('div');
  toast.className = 'toast-notification';
  toast.innerHTML = `
    <div style="display:flex; align-items:center; gap:10px; background: rgba(18,20,28,0.95); border: 1px solid #ff2a39; color: #fff; padding: 14px 24px; border-radius: 8px; box-shadow: 0 10px 30px rgba(0,0,0,0.8); position: fixed; bottom: 30px; right: 30px; z-index: 9999; font-weight:500;">
      <span style="color:#ff2a39;">✓</span> ${message}
    </div>
  `;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.5s ease';
    setTimeout(() => toast.remove(), 500);
  }, 3500);
}

// Attach toast to CTAs
document.addEventListener('click', (e) => {
  if (e.target.closest('.book-test-drive-btn')) {
    e.preventDefault();
    showToast('Test drive booking initiated! Our Audi specialist will contact you.');
  }
});
