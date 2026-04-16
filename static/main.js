document.addEventListener('DOMContentLoaded', () => {
    fetchData();
    // Poll every 10 seconds for updates
    setInterval(fetchData, 10000);
});

async function fetchData() {
    try {
        const response = await fetch('/api/progress');
        const result = await response.json();
        
        if (result.status === 'success') {
            renderDashboard(result.data);
        } else {
            console.error('Error fetching data:', result.message);
        }
    } catch (error) {
        console.error('Network error:', error);
    }
}

function renderDashboard(data) {
    const grid = document.getElementById('staff-grid');
    const template = document.getElementById('card-template');
    
    // Clear existing grid unless we want to do smart updates
    grid.innerHTML = '';
    
    let totalProgress = 0;
    
    data.forEach(item => {
        // Fallbacks for data
        const name = item['Employee Name'] || 'Unknown';
        const task = item['Task Name'] || 'No Task Assigned';
        const progress = item['Progress %'] || 0;
        const status = item['Status'] || 'Pending';
        
        totalProgress += progress;
        
        const clone = template.content.cloneNode(true);
        
        // Populate Data
        clone.querySelector('.employee-name').textContent = name;
        
        const statusBadge = clone.querySelector('.status-badge');
        statusBadge.textContent = status;
        statusBadge.setAttribute('data-status', status);
        
        clone.querySelector('.task-name').textContent = task;
        clone.querySelector('.progress-value').textContent = `${progress}%`;
        
        // Setup Progress Bar Fill dynamically
        const fill = clone.querySelector('.progress-fill');
        
        // Force reflow animation
        setTimeout(() => {
            fill.style.width = `${progress}%`;
            
            // Optional: change color based on progress
            if(progress === 100) {
                fill.style.background = 'var(--success-color)';
                fill.style.boxShadow = '0 0 10px rgba(16, 185, 129, 0.5)';
            }
        }, 50); // slight delay so the CSS transition kicks in when added to DOM
        
        grid.appendChild(clone);
    });
    
    // Update Overall Progress
    const avgProgress = data.length ? Math.round(totalProgress / data.length) : 0;
    document.getElementById('overall-percentage').textContent = `${avgProgress}%`;
    
    // SVG Circle Math
    // The stroke-dasharray is basically `<filled>, 100` where the total length is approx 100 because of the path we drew
    const circle = document.getElementById('overall-chart-val');
    setTimeout(() => {
        circle.setAttribute('stroke-dasharray', `${avgProgress}, 100`);
    }, 50);
}
