import { Link, Routes, Route } from "react-router-dom";
import "./App.css";
import ManPowerPage from "./pages/ManPowerPage";

function App() {
  return (
    <>
      <h1>Costing Dashboard</h1>
      <p>เลือกเมนูด้านล่างเพื่อเข้าสู่หน้าที่ต้องการ</p>

      {/* การ์ดลิงก์ไปแต่ละหน้า */}
      <div className="card-container">
        <Link to="/manpower" className="card-link">
          <div className="card">
            <h2>Manpower Upload</h2>
          </div>
        </Link>
      </div>

      {/* เส้นทางแต่ละหน้า */}
      <Routes>
        <Route path="/manpower" element={<ManPowerPage />} />
      </Routes>
    </>
  );
}

export default App;
