import streamlit as st
import streamlit.components.v1 as components
from ui.pages.alltrans_page import run_mvr_all_trans
from ui.pages.hdvi_page import run_hdvi_mvr
from ui.pages.riscom_page import run_riscom_mvr
from ui.pages.riscom_renewal_page import run_riscom_renewal_mvr


def run_mvr_summary():

    components.html("""
<!DOCTYPE html>
<html>
<head>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;600&family=Inter:wght@300;400;500&display=swap" rel="stylesheet">
<style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: transparent; overflow: hidden; }

    .scene {
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        justify-content: center;
        min-height: 130px;
        padding: 1.75rem 0.5rem 1.5rem 0.25rem;
        overflow: hidden;
    }

    .deep-glow {
        position: absolute;
        width: 480px; height: 200px;
        top: 50%; left: 30%;
        transform: translate(-50%, -50%);
        background: radial-gradient(ellipse at center,
            rgba(80, 100, 220, 0.06) 0%,
            rgba(50, 70, 180, 0.03) 45%,
            transparent 70%);
        pointer-events: none;
        filter: blur(20px);
        animation: glowBreathe 9s ease-in-out infinite alternate;
    }

    #starCanvas {
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;
    }

    .header-content {
        position: relative;
        z-index: 2;
    }

    .eyebrow {
        font-family: 'Inter', sans-serif;
        font-size: 0.62rem;
        font-weight: 500;
        letter-spacing: 0.2em;
        text-transform: uppercase;
        color: rgba(99, 118, 255, 0.85);
        margin-bottom: 0.55rem;
        opacity: 0;
        animation: fadeUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.2s;
    }

    .page-title {
        font-family: 'Cormorant Garamond', Georgia, serif;
        font-size: clamp(1.7rem, 4vw, 2.2rem);
        font-weight: 300;
        letter-spacing: 0.01em;
        line-height: 1.15;
        color: #E4E8F5;
        opacity: 0;
        animation: fadeUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.4s;
    }

    .page-title em {
        font-style: italic;
        font-weight: 600;
        background: linear-gradient(120deg, #9BAEFF 0%, #C8CFFF 45%, #7DD6EA 100%);
        background-size: 220% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: shimmer 6s linear 1.5s infinite;
    }

    .page-subtitle {
        font-family: 'Inter', sans-serif;
        font-size: 0.8rem;
        font-weight: 300;
        color: rgba(130, 142, 172, 0.75);
        letter-spacing: 0.03em;
        margin-top: 0.5rem;
        opacity: 0;
        animation: fadeUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.65s;
    }

    .rule-bottom {
        position: relative;
        z-index: 2;
        width: 0;
        height: 1px;
        background: linear-gradient(90deg, rgba(99,118,255,0.4), rgba(6,182,212,0.35), transparent);
        margin-top: 1.4rem;
        animation: expandRule 1.4s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.9s;
    }

    @keyframes fadeUp {
        from { opacity: 0; transform: translateY(14px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes expandRule {
        from { width: 0; }
        to   { width: 320px; }
    }
    @keyframes shimmer {
        0%   { background-position: 0% center; }
        100% { background-position: 220% center; }
    }
    @keyframes glowBreathe {
        from { opacity: 0.7; transform: translate(-50%, -50%) scale(1); }
        to   { opacity: 1;   transform: translate(-50%, -50%) scale(1.1); }
    }
</style>
</head>
<body>
<div class="scene">
    <div class="deep-glow"></div>
    <canvas id="starCanvas"></canvas>

    <div class="header-content">
        <p class="eyebrow">Motor Vehicle Records</p>
        <h1 class="page-title">MVR Summary Reports</h1>
        <p class="page-subtitle">Process and analyze Motor Vehicle Records from multiple sources</p>
        <div class="rule-bottom"></div>
    </div>
</div>

<script>
    const canvas = document.getElementById('starCanvas');
    const ctx = canvas.getContext('2d');

    function resize() {
        canvas.width  = canvas.offsetWidth  || 800;
        canvas.height = canvas.offsetHeight || 130;
    }
    resize();

    const W = canvas.width, H = canvas.height;
    const stars = [];

    for (let i = 0; i < 90; i++) {
        stars.push({
            x: Math.random() * W, y: Math.random() * H,
            r: Math.random() * 0.55 + 0.15,
            baseAlpha: Math.random() * 0.16 + 0.05,
            twinkleSpeed: Math.random() * 0.004 + 0.001,
            twinkleOffset: Math.random() * Math.PI * 2,
            tier: 1
        });
    }
    for (let i = 0; i < 25; i++) {
        stars.push({
            x: Math.random() * W, y: Math.random() * H,
            r: Math.random() * 0.75 + 0.5,
            baseAlpha: Math.random() * 0.2 + 0.09,
            twinkleSpeed: Math.random() * 0.003 + 0.001,
            twinkleOffset: Math.random() * Math.PI * 2,
            tier: 2
        });
    }
    for (let i = 0; i < 5; i++) {
        stars.push({
            x: Math.random() * W, y: Math.random() * H,
            r: Math.random() * 0.9 + 0.9,
            baseAlpha: Math.random() * 0.25 + 0.1,
            twinkleSpeed: Math.random() * 0.002 + 0.0005,
            twinkleOffset: Math.random() * Math.PI * 2,
            tier: 3
        });
    }

    let t = 0;
    function draw() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        t += 0.016;

        for (const s of stars) {
            const alpha = s.baseAlpha * (0.6 + 0.4 * Math.sin(t * s.twinkleSpeed * 60 + s.twinkleOffset));

            if (s.tier === 3) {
                const grd = ctx.createRadialGradient(s.x, s.y, 0, s.x, s.y, s.r * 3.5);
                grd.addColorStop(0, `rgba(190, 200, 255, ${alpha})`);
                grd.addColorStop(1, `rgba(190, 200, 255, 0)`);
                ctx.beginPath();
                ctx.arc(s.x, s.y, s.r * 3.5, 0, Math.PI * 2);
                ctx.fillStyle = grd;
                ctx.fill();
            }

            ctx.beginPath();
            ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
            const col = s.tier === 1 ? '210, 215, 240' : '200, 210, 255';
            ctx.fillStyle = `rgba(${col}, ${alpha})`;
            ctx.fill();
        }

        requestAnimationFrame(draw);
    }
    draw();
</script>
</body>
</html>
""", height=150)

    st.markdown("")

    col1, col2, col3 = st.columns([2, 3, 2])
    with col2:
        report_type = st.selectbox(
            "Choose MVR Report:",
            options=["HDVI", "NB-Riscom", "Alltrans","Renewal-Riscom"],
            label_visibility="collapsed",
            key="mvr_report_selector"
        )

    st.markdown("")

    try:
        if report_type == "HDVI":
            run_hdvi_mvr()
        elif report_type == "NB-Riscom":
            run_riscom_mvr()
        elif report_type == "Alltrans":
            run_mvr_all_trans()
        elif report_type == "Renewal-Riscom":
            run_riscom_renewal_mvr()

    except Exception as e:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.error(f"Error loading report: {str(e)}")
        with col2:
            if st.button("Show Details", key="mvr_error_details"):
                with st.expander("Technical Details"):
                    import traceback
                    st.code(traceback.format_exc())