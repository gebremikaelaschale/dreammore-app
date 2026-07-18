import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import * as XLSX from 'xlsx';
import App from './App';

test('renders the main portal heading', () => {
  render(<App />);
  expect(screen.getByText(/DreamMore Enrollment Portal/i)).toBeInTheDocument();
});

test('shows uploaded student preview and course options after parsing a spreadsheet', async () => {
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Name', 'Email', 'Course'],
    ['Alice Johnson', 'alice@example.com', 'Python'],
    ['Bob Smith', 'bob@example.com', 'Design']
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Students');
  const fileBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
  const file = new File([fileBuffer], 'students.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

  render(<App />);
  fireEvent.change(screen.getByLabelText(/upload excel file/i), { target: { files: [file] } });

  await waitFor(() => {
    expect(screen.getByText(/Alice Johnson/i)).toBeInTheDocument();
  });

  expect(screen.getByRole('combobox', { name: /course/i })).toBeInTheDocument();
  expect(screen.getByText(/Bob Smith/i)).toBeInTheDocument();
});
