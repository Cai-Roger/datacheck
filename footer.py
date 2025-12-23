import streamlit.components.v1 as components

def render_footer():
    components.html(
        """
        <style>
          [data-testid="stMainBlockContainer"] {
              padding-bottom: 60px;
          }

          .app-footer {
              position: fixed;
              left: 0;
              bottom: 0;
              width: 100%;
              background-color: #f5f6f7;
              color: #555;
              text-align: center;
              padding: 10px 0;
              font-size: 13px;
              border-top: 1px solid #e0e0e0;
              z-index: 9999;
          }
        </style>

        <div class="app-footer">
            © 2025 Roger & Andy ｜ Excel 比對程式 ｜ V2.1.2
        </div>
        """,
        height=0
    )
