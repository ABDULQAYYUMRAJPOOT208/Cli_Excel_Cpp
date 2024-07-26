#include "Header.h"
#include <string>
#include <conio.h>

template <typename T>
void takeInput(Node<T>* curr);
void firstMenu();
rangeStart rStartInput();
rangeEnd rEndInput();
string option();
int main()
{
	string input, fileName = "excel.txt";
	excel<string> ex;
	if(!ex.loadExcelFromFile(ex, fileName));
	{
		ex.saveToFile(fileName);
	}
	ex.printExcel();
	firstMenu();

	while (true)
	{
		if (GetAsyncKeyState(VK_UP) and ex.current->up != nullptr)
		{
			ex.current = ex.current->up;
			ex.printExcel();
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_DOWN) and ex.current->down != nullptr)
		{
			ex.current = ex.current->down;
			ex.printExcel();
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_LEFT) and ex.current->left != nullptr)
		{
			ex.current = ex.current->left;
			ex.printExcel();
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_RIGHT) and ex.current->right != nullptr)
		{
			ex.current = ex.current->right;
			ex.printExcel();
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_CONTROL))
		{
			input = option();
			if (input == "2")
			{
				ex.createGRow();
			}
			else if (input == "3")
			{
				ex.createRowUp();
			}
			else if (input == "4")
			{
				ex.insertColumnToRight();
			}
			else if (input == "5")
			{
				ex.insertColumnToLeft();
			}
			else if (input == "6")
			{
				ex.insertCellByLeftShift();
			}
			else if (input == "7")
			{
				ex.insertCellByRightShift();
			}
			else if (input == "8")
			{
				ex.insertCellByUpShift();
			}
			else if (input == "9")
			{
				ex.insertCellByDownShift();
			}
			else if (input == "10")
			{
				ex.deleteCellByLeftShift();
			}
			else if (input == "11")
			{
				ex.deleteCellByUpShift(ex.current);
			}
			else if (input == "12")
			{
				ex.deleteColumn();
			}
			else if (input == "13")
			{
				ex.deleteRowCurrent();
			}
			else if (input == "14")
			{
				ex.ClearRow();
			}
			else if (input == "15")
			{
				ex.ClearColumn();
			}
			else if (input == "16")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				cout << "Sum is: " << ex.GetRangeSum(start, end) << endl;
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "17")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				cout << "Average is: " << ex.GetRangeAverage(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "18")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				cout << "Selectd Cells are: " << ex.getSelectedCount(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "19")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				cout << "Minimum is: " << ex.GetRangeMinimum(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "20")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				cout << "Maximum is: " << ex.GetRangeMaximum(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "21")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				ex.RangeCopySelect(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "22")
			{
				ex.paste();
				cout << "Press any key to continue...";
				_getch();
			}
			else if (input == "23")
			{
				rangeStart start = rStartInput();
				rangeEnd end = rEndInput();
				ex.RangeCutSelect(start, end);
				cout << "Press any key to continue...";
				_getch();
			}
			ex.saveToFile(fileName);
			ex.printExcel();
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_SHIFT))
		{
			takeInput(ex.current);
			ex.printExcel();
			ex.saveToFile(fileName);
			firstMenu();
		}
		else if (GetAsyncKeyState(VK_ESCAPE))
		{
			break;
		}

		Sleep(50);
	}
	return 0;
}
template <typename T>
void takeInput(Node<T>* curr)
{
	string str;
	cout << "Enter Data: ";
	getline(cin, str);
	if (str == "$$$")
	{
		cout << "This character is not allowed..." << endl;
		cout << "Press any key to continue...";
		_getch();
	}
	else
	{
		size_t found = str.find(',');
		if (found != string::npos)
		{
			cout << "This character Comma is not allowed..." << endl;
			cout << "Press any key to continue...";
			_getch();
		}
		else
		{
			curr->data = str;
			str.clear();
		}
	}
};
void firstMenu()
{
	cout << "Press Shift to Enter Data in current cell" << endl;
	cout << "Press Control For Other Options" << endl;
}
string option()
{
	string str;
	cout << "Enter 1 to Exit" << endl;
	cout << "Enter 2 to insert a Row below current Cell" << endl;
	cout << "Enter 3 to insert a Row Above the current Cell" << endl;
	cout << "Enter 4 to insert a Column Right to the current Cell" << endl;
	cout << "Enter 5 to insert a Column Left to the current Cell" << endl;
	cout << "Enter 6 to insert a Cell by left shift" << endl;
	cout << "Enter 7 to insert a Cell by Right shift" << endl;
	cout << "Enter 8 to insert a Cell by Up shift" << endl;
	cout << "Enter 9 to insert a Cell by Down shift" << endl;
	cout << "Enter 10 to delete current Cell by left shift" << endl;
	cout << "Enter 11 to delete current Cell by Up shift" << endl;
	cout << "Enter 12 to delete current Column " << endl;
	cout << "Enter 13 to delete current Row " << endl;
	cout << "Enter 14 to Clear current Row " << endl;
	cout << "Enter 15 to Clear current Column " << endl;
	cout << "Enter 16 for sum Function " << endl;
	cout << "Enter 17 for average Function " << endl;
	cout << "Enter 18 for selected count Function " << endl;
	cout << "Enter 19 for range minimum Function " << endl;
	cout << "Enter 20 for range maximum Function " << endl;
	cout << "Enter 21 for copying Function " << endl;
	cout << "Enter 22 for paste Function " << endl;
	cout << "Enter 23 for range Cut Function " << endl;
	cout << "Select Options : ";
	cin >> str;
	return str;
}

rangeStart rStartInput()
{
	char r, c;
	cout << "Enter starting row: ";
	cin >> r;
	cout << "Enter starting column: ";
	cin >> c;
	if (isdigit(c) and isdigit(r))
	{
		int row = r - '0';
		int col = c - '0';
		return rangeStart(row, col);
	}
	else
	{
		return rangeStart(0, 0);
	}
}
rangeEnd rEndInput()
{
	char r, c;
	cout << "Enter ending row: ";
	cin >> r;
	cout << "Enter ending column: ";
	cin >> c;
	if (isdigit(c) and isdigit(r))
	{
		int row = r - '0';
		int col = c - '0';
		return rangeEnd(row, col);
	}
	else
	{
		return rangeEnd(0, 0);
	}
}
