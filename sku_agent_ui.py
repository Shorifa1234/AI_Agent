"""
SKU Agent Chat UI
Run: streamlit run sku_agent_ui.py
"""

import streamlit as st
import subprocess
import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Vendor registry ──────────────────────────────────────────────────────────
VENDORS = {
    "crystoramalights": {
        "name": "Crystorama Lights (crystoramalightinglights.com)",
        "keywords": ["crystorama light", "crylights", "crystoramalight"],
        "steps": {
            "step1": {"script": "CrystoramaLights/crylights_step1.py", "dir": "CrystoramaLights"},
            "step2": {"script": "CrystoramaLights/crylights_step2.py", "dir": "CrystoramaLights"},
        },
    },
    "crystorama": {
        "name": "Crystorama (crystorama.com)",
        "keywords": ["crystorama.com", "crystorama"],
        "steps": {
            "step1": {"script": "Crystorama/crystorama_step1.py", "dir": "Crystorama"},
            "step2": {"script": "Crystorama/crystorama_step2.py", "dir": "Crystorama"},
        },
    },
    "vanguard": {
        "name": "Vanguard Designs",
        "keywords": ["vanguard designs", "vanguard"],
        "steps": {
            "step1": {"script": "Vanguard/vanguard_step1.py", "dir": "Vanguard"},
            "step2": {"script": "Vanguard/vanguard_step2.py", "dir": "Vanguard"},
        },
    },
    "brownstone": {
        "name": "Brownstone",
        "keywords": ["brownstone"],
        "steps": {
            "step1": {"script": "Brownstone/brownstone_step1.py", "dir": "Brownstone"},
            "step2": {"script": "Brownstone/brownstone_step2.py", "dir": "Brownstone"},
        },
    },
    "invisible": {
        "name": "Invisible Collection",
        "keywords": ["invisible collection", "invisible"],
        "steps": {
            "step1": {"script": "Invisible Collection/invisible_step1.py", "dir": "Invisible Collection"},
            "step2": {"script": "Invisible Collection/invisible_step2.py", "dir": "Invisible Collection"},
        },
    },
    "wesleyhall": {
        "name": "Wesley Hall",
        "keywords": ["wesley hall", "wesleyhall", "wesley"],
        "steps": {
            "step1": {"script": "Wesley Hall/step1.py", "dir": "Wesley Hall"},
            "step2": {"script": "Wesley Hall/step2.py", "dir": "Wesley Hall"},
        },
    },
    "paulferrante": {
        "name": "Paul Ferrante",
        "keywords": ["paul ferrante", "ferrante"],
        "steps": {
            "step1": {"script": "Paul Ferrante/listpage.py",    "dir": "Paul Ferrante"},
            "step2": {"script": "Paul Ferrante/detailspage.py", "dir": "Paul Ferrante"},
            "step3": {"script": "Paul Ferrante/step3.py",       "dir": "Paul Ferrante"},
        },
    },
    "blackmancruz": {
        "name": "Blackman Cruz",
        "keywords": ["blackman cruz", "blackman", "blackmancruz"],
        "steps": {
            "step1": {"script": "BlackmanCruz/blackman_step1.py", "dir": "BlackmanCruz"},
            "step2": {"script": "BlackmanCruz/blackman_step2.py", "dir": "BlackmanCruz"},
        },
    },
    "gabby": {
        "name": "Gabby",
        "keywords": ["gabby"],
        "steps": {
            "step1": {"script": "Gabby/gabby_step1.py", "dir": "Gabby"},
            "step2": {"script": "Gabby/gabby_step2.py", "dir": "Gabby"},
        },
    },
    "worldsaway": {
        "name": "Worlds Away",
        "keywords": ["worlds away", "worldsaway"],
        "steps": {
            "step1": {"script": "WorldsAway/worldsaway_step1.py", "dir": "WorldsAway"},
            "step2": {"script": "WorldsAway/worldsaway_step2.py", "dir": "WorldsAway"},
        },
    },
}

STEP_LABELS = {
    "step1": "Step 1 — Product List",
    "step2": "Step 2 — Detail Scraper",
    "step3": "Step 3 — Final Export",
}


# ── Parse user instruction ────────────────────────────────────────────────────
def parse_instruction(text: str):
    tl = text.lower()

    # Detect vendor (longest keyword match first to avoid partial collisions)
    vendor_key = None
    best_len = 0
    for k, v in VENDORS.items():
        for kw in v["keywords"]:
            if kw in tl and len(kw) > best_len:
                vendor_key = k
                best_len = len(kw)

    # Detect step
    if "step1" in tl or "step 1" in tl:
        step = "step1"
    elif "step3" in tl or "step 3" in tl:
        step = "step3"
    elif "step2" in tl or "step 2" in tl:
        step = "step2"
    else:
        step = "all"

    return vendor_key, step


# ── Run a single script, stream output ───────────────────────────────────────
def run_script(script_rel: str, work_dir_rel: str, log_area) -> bool:
    script_path = os.path.join(BASE_DIR, script_rel)
    work_dir    = os.path.join(BASE_DIR, work_dir_rel)

    if not os.path.exists(script_path):
        log_area.error(f"Script not found: {script_path}")
        return False

    proc = subprocess.Popen(
        [sys.executable, "-u", script_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=work_dir,
        bufsize=1,
    )

    lines = []
    for raw in proc.stdout:
        line = raw.rstrip()
        lines.append(line)
        log_area.code("\n".join(lines[-30:]))  # rolling 30-line window

    proc.wait()
    return proc.returncode == 0


# ── Streamlit app ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="SKU Agent", page_icon="🤖", layout="wide")
st.title("🤖 SKU Scraping Agent")

# Sidebar — vendor list
with st.sidebar:
    st.subheader("Available Vendors")
    for v in VENDORS.values():
        st.markdown(f"• {v['name']}")
    st.divider()
    st.caption("Example instructions:")
    st.code("Crystorama step2 run koro")
    st.code("Vanguard full scrape koro")
    st.code("Wesley Hall step1 run koro")
    st.code("Brownstone both steps run koro")

# Session state
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "👋 Assalamu Alaikum! Vendor-er naam aar step likhun.\n\n"
                "**Example:**\n"
                "- `Crystorama Lights step2 run koro`\n"
                "- `Vanguard scrape koro` (step1 + step2 both)\n"
                "- `Wesley Hall step1 run koro`"
            ),
        }
    ]

# Render chat history
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Chat input
if prompt := st.chat_input("Vendor instruction likhun..."):
    # Show user message
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Parse
    vendor_key, step = parse_instruction(prompt)

    with st.chat_message("assistant"):
        if vendor_key is None:
            reply = (
                "❌ Vendor detect korte parini.\n\n"
                "Please vendor-er naam likhunn, e.g.:\n"
                "`Crystorama step2 run koro`"
            )
            st.markdown(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})

        else:
            vendor = VENDORS[vendor_key]
            available_steps = list(vendor["steps"].keys())

            # Decide which steps to run
            if step == "all":
                steps_to_run = available_steps
            elif step in vendor["steps"]:
                steps_to_run = [step]
            else:
                reply = (
                    f"❌ **{vendor['name']}** এর জন্য `{step}` নেই।\n"
                    f"Available steps: {', '.join(available_steps)}"
                )
                st.markdown(reply)
                st.session_state.messages.append({"role": "assistant", "content": reply})
                steps_to_run = []

            all_ok = True
            for s in steps_to_run:
                info = vendor["steps"][s]
                label = STEP_LABELS.get(s, s)
                st.markdown(f"▶️ **{vendor['name']}** — `{label}` শুরু হচ্ছে...")
                log_area = st.empty()
                ok = run_script(info["script"], info["dir"], log_area)
                if ok:
                    st.success(f"✅ `{label}` সম্পন্ন হয়েছে!")
                else:
                    st.error(f"❌ `{label}` এ সমস্যা হয়েছে।")
                    all_ok = False
                    break

            if steps_to_run:
                final = (
                    f"✅ **{vendor['name']}** — সব steps সম্পন্ন!"
                    if all_ok
                    else f"⚠️ **{vendor['name']}** — কিছু steps সম্পন্ন হয়নি।"
                )
                st.session_state.messages.append({"role": "assistant", "content": final})
