export default function Card({ title, children }) {
  return (
    <div className="bg-white shadow-md rounded-2xl p-5 border border-blue-100">
      {title && <h2 className="text-lg font-semibold text-blue-600 mb-3 border-b pb-2">{title}</h2>}
      {children}
    </div>
  );
}
