#pragma once
#include <iostream>
#include <windows.h>
#include <vector>
#include <fstream>
#include <sstream>
#include <string>
using namespace std;

class rangeStart
{
public:
	int rowStart, colStart;

	rangeStart(int s, int e) : rowStart(s), colStart(e) {}
};

class rangeEnd
{
public:
	int rowEnd, colEnd;
	rangeEnd(int re, int ce) : rowEnd(re), colEnd(ce) {}
};
template <typename T>
class Node
{
public:
	Node* up;
	Node* down;
	Node* right;
	Node* left;
	// T data;
	T data;

	Node(T val) : data(val), up(nullptr), down(nullptr), right(nullptr), left(nullptr) {}
};
template <typename T>
class excel
{
public:
	Node<T>* ex;
	Node<T>* head;
	Node<T>* current;
	int rows, columns;
	vector<T> pasteVecter;
	excel() : rows(0), columns(0), ex(nullptr), head(nullptr), current(nullptr)
	{
		createFirstRow();
		for (int i = 2; i <= 5; i++)
		{
			createGRow();
		}
		rows = columns = 5;
	}

	void createFirstRow()
	{
		for (int col = 1; col <= 5; col++)
		{
			Node<T>* nnode = new Node<T>("NIL");
			if (ex == nullptr)
			{
				ex = nnode;
				head = nnode;
				current = nnode;
			}
			else
			{
				current->right = nnode;
				nnode->left = current;
				current = current->right;
			}
		}
		rows++;
	}
	void createGRow()
	{
		Node<T>* temp = this->current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		Node<T>* nfirst = new Node<T>("NIL");

		if (temp->down == nullptr)
		{
			temp->down = nfirst;
			nfirst->up = temp;
		}
		else
		{
			temp->down->up = nfirst;
			nfirst->down = temp->down;
			nfirst->up = temp;
			temp->down = nfirst;
		}
		temp = temp->right;
		Node<T>* temp2 = nfirst;
		while (temp != nullptr)
		{

			Node<T>* nnode = new Node<T>("NIL");
			if (temp->down == nullptr)
			{
				temp->down = nnode;
				nnode->up = temp;
				temp2->right = nnode;
				nnode->left = temp2;
			}
			else
			{
				temp->down->up = nnode;
				nnode->down = temp->down;
				nnode->up = temp;
				temp->down = nnode;
				nnode->left = temp2;
				temp2->right = nnode;
			}
			temp2 = temp2->right;
			temp = temp->right;
			current = temp2;
		}
		rows++;
	}

	// Iterator
	class iiterator
	{
	public:
		Node<T>* iter;
		iiterator(Node<T>* i)
		{
			iter = i;
		}
		iiterator operator++(int)
		{

			return iter = iter->right;
		}
		iiterator operator--(int)
		{
			return iter = iter->left;
		}
		iiterator operator--()
		{
			return iter = iter->up;
		}
		iiterator operator++()
		{
			return iter = iter->down;
		}
	};

	// ROW Above
	void createRowUp()
	{
		Node<T>* temp = this->current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		Node<T>* nfirst = new Node<T>("NIL");

		if (temp->up == nullptr)
		{
			temp->up = nfirst;
			nfirst->down = temp;
			head = nfirst;
		}
		else
		{
			temp->up->down = nfirst;
			nfirst->up = temp->up;
			nfirst->down = temp;
			temp->up = nfirst;
		}
		temp = temp->right;
		Node<T>* temp2 = nfirst;

		while (temp != nullptr)
		{

			Node<T>* nnode = new Node<T>("NIL");
			if (temp->up == nullptr)
			{
				temp->up = nnode;
				nnode->down = temp;
				temp2->right = nnode;
				nnode->left = temp2;
			}
			else
			{
				temp->up->down = nnode;
				nnode->up = temp->up;
				nnode->down = temp;
				temp->up = nnode;
				nnode->left = temp2;
				temp2->right = nnode;
			}
			temp2 = temp2->right;
			temp = temp->right;
			current = temp2;
		}
		rows++;
	}
	void InserRowAbove()
	{
		Node<T>* nnode = current;
		while (nnode->left != nullptr)
		{
			nnode = nnode->left;
		}

		while (nnode != nullptr)
		{
			Node<T>* newNode = new Node<T>("NIL");
			newNode->up = nnode->up;
			if (nnode->up != nullptr)
			{
				nnode->up->down = newNode;
			}
			newNode->down = nnode;
			nnode->up = newNode;
			if (nnode->left != nullptr)
			{
				newNode->left = nnode->left->up;
				if (nnode->left->up != nullptr)
				{
					nnode->left->up->right = newNode;
				}
			}
			nnode = nnode->right;
		}
		if (head->up != nullptr)
		{
			head = head->up;
		}
		rows++;
	}
	void InserRowBelow()
	{
		Node<T>* nnode = current;
		while (nnode->left != nullptr)
		{
			nnode = nnode->left;
		}

		while (nnode != nullptr)
		{
			Node<T>* newNode = new Node<T>("NIL");
			newNode->down = nnode->down;
			if (nnode->down != nullptr)
			{
				nnode->down->up = newNode;
			}
			newNode->up = nnode;
			nnode->down = newNode;
			if (nnode->left != nullptr)
			{
				newNode->left = nnode->left->down;
				if (nnode->left->down != nullptr)
				{
					nnode->left->down->right = newNode;
				}
			}
			nnode = nnode->right;
		}
		rows++;
	}
	void printExcel()
	{
		char c = 240;
		system("cls");
		HANDLE H = GetStdHandle(STD_OUTPUT_HANDLE);
		Node<T>* temp = head;
		while (temp != nullptr)
		{
			Node<T>* current = temp;
			while (current != nullptr)
			{
				cout << "+" << c << c << c << c << c << c;
				current = current->right;
			}
			cout << "+" << endl;

			current = temp;
			while (current != nullptr)
			{
				if (current == this->current)
				{

					SetConsoleTextAttribute(H, 2);
					cout << "|";
					if (current->data.length() < 4)
					{
						for (int i = current->data.length(); i <= 4; i++)
						{
							cout << " ";
						}
						cout << current->data << " ";
					}
					else if (this->current->data.length() == 4)
					{
						cout << " ";
						cout << current->data << " ";
					}
					else
					{
						string sub = current->data.substr(0, 4);
						cout << " " << sub << " ";
					}
					SetConsoleTextAttribute(H, 7);
				}
				else
				{
					cout << "|";
					if (current->data.length() < 4)
					{
						for (int i = current->data.length(); i <= 4; i++)
						{
							cout << " ";
						}
						cout << current->data << " ";
					}
					else if (this->current->data.length() == 4)
					{
						cout << " ";
						cout << current->data << " ";
					}
					else
					{
						string sub = current->data.substr(0, 4);
						cout << " " << sub << " ";
					}
				}
				current = current->right;
			}
			cout << "|" << endl;

			// cout << "|      |";
			temp = temp->down;
		}

		// Print the bottom border
		Node<T>* current = head;
		while (current != nullptr)
		{
			cout << "+" << c << c << c << c << c << c;
			current = current->right;
		}
		cout << "+" << endl;
	}

	// columns insert
	void insertColumnToRight()
	{
		Node<T>* temp = this->current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		Node<T>* nfirst = new Node<T>("NIL");
		if (temp->right == nullptr)
		{
			temp->right = nfirst;
			nfirst->left = temp;
		}
		else
		{
			temp->right->left = nfirst;
			nfirst->right = temp->right;
			nfirst->left = temp;
			temp->right = nfirst;
		}
		temp = temp->down;
		Node<T>* temp2 = nfirst;

		while (temp != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			if (temp->right == nullptr)
			{
				temp->right = nnode;
				nnode->left = temp;
				nnode->up = temp2;
				temp2->down = nnode;
			}
			else
			{
				temp->right->left = nnode;
				nnode->right = temp->right;
				nnode->left = temp;
				temp->right = nnode;
				temp->up = temp2;
				temp2->down = nnode;
			}
			temp2 = temp2->down;
			temp = temp->down;
			current = temp2;
		}
		this->columns++;
	}
	void insertColumnToLeft()
	{
		Node<T>* temp = this->current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		Node<T>* nfirst = new Node<T>("NIL");
		if (temp == head)
		{
			head = nfirst;
		}
		if (temp->left == nullptr)
		{
			temp->left = nfirst;
			nfirst->right = temp;
		}
		else
		{
			temp->left->right = nfirst;
			nfirst->left = temp->left;
			nfirst->right = temp;
			temp->left = nfirst;
		}
		temp = temp->down;
		Node<T>* temp2 = nfirst;

		while (temp != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			if (temp->left == nullptr)
			{
				temp->left = nnode;
				nnode->right = temp;
				nnode->up = temp2;
				temp2->down = nnode;
			}
			else
			{
				temp->left->right = nnode;
				nnode->left = temp->left;
				nnode->right = temp;
				temp->left = nnode;
				nnode->up = temp2;
				temp2->down = nnode;
			}
			temp2 = temp2->down;
			temp = temp->down;
		}
		this->columns++;
	}
	// Helping Function starts
	void insertFirstColumn()
	{
		Node<T>* temp = head;
		Node<T>* nfirst = new Node<T>("NIL");
		temp->left = nfirst;
		nfirst->right = temp;
		head = nfirst;
		Node<T>* chase = nfirst;
		temp = temp->down;
		while (temp != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			nnode->up = chase;
			chase->down = nnode;
			nnode->right = temp;
			temp->left = nnode;
			temp = temp->down;
			chase = chase->down;
		}
	}
	void insertLastColumn()
	{
		Node<T>* last = head;
		while (last->right != nullptr)
		{
			last = last->right;
		}

		Node<T>* nfirst = new Node<T>("NIL");
		last->right = nfirst;
		nfirst->left = last;
		last = last->down;
		Node<T>* chase = nfirst;
		while (last != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			nnode->up = chase;
			chase->down = nnode;
			last->right = nnode;
			nnode->left = last;

			last = last->down;
			chase = chase->down;
		}
	}
	void insertFirstRow()
	{
		Node<T>* temp = head;
		Node<T>* nfirst = new Node<T>("NIL");
		temp->up = nfirst;
		nfirst->down = temp;
		head = nfirst;
		Node<T>* chase = nfirst;
		temp = temp->right;
		while (temp != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			nnode->left = chase;
			chase->right = nnode;
			temp->up = nnode;
			nnode->down = temp;

			temp = temp->right;
			chase = chase->right;
		}
	}
	void insertLastRow()
	{
		Node<T>* last = head;
		while (last->down != nullptr)
		{
			last = last->down;
		}
		Node<T>* nfirst = new Node<T>("NIL");
		last->down = nfirst;
		nfirst->up = last;
		last = last->right;
		Node<T>* chase = nfirst;
		while (last != nullptr)
		{
			Node<T>* nnode = new Node<T>("NIL");
			nnode->left = chase;
			chase->right = nnode;
			nnode->up = last;
			last->down = nnode;

			last = last->right;
			chase = chase->right;
		}
	}
	// Helping Function Ends

	void insertCellByUpShift()
	{
		insertFirstRow();
		Node<T>* first = this->current;
		while (first->up != nullptr)
		{
			first = first->up;
		}

		Node<T>* chase = first;
		while (chase != current)
		{
			chase->data = chase->down->data;
			chase = chase->down;
		}
		this->current->data = " ";
		this->rows++;
	}
	void insertCellByRightShift()
	{
		insertLastColumn();
		Node<T>* last = this->current;
		while (last->right != nullptr)
		{
			last = last->right;
		}
		Node<T>* chase = last;
		while (chase != current)
		{
			chase->data = chase->left->data;
			chase = chase->left;
		}
		this->current->data = " ";
		this->columns++;
	}
	void insertCellByLeftShift()
	{
		insertFirstColumn();
		Node<T>* first = this->current;
		while (first->left != nullptr)
		{
			first = first->left;
		}
		Node<T>* chase = first;
		while (chase != current)
		{
			chase->data = chase->right->data;
			chase = chase->right;
		}
		this->current->data = " ";
		this->columns++;
	}
	void insertCellByDownShift()
	{
		insertLastRow();
		Node<T>* last = this->current;
		while (last->down != nullptr)
		{
			last = last->down;
		}
		Node<T>* chase = last;
		while (chase != current)
		{
			chase->data = chase->up->data;
			chase = chase->up;
		}
		this->current->data = " ";
		this->rows++;
	}
	void deleteCellByLeftShift()
	{
		if (this->current->right == nullptr)
		{
			this->current = nullptr;
			this->current->data = "";
		}
		else
		{
			Node<T>* temp = this->current;
			while (temp->right->right != nullptr)
			{
				temp->data = temp->right->data;
				temp = temp->right;
			}

			temp->right->data = "";
		}
	}
	// THis is incomplete......
	void deleteCellByUpShift(Node<T>* c)
	{
		if (c->up == nullptr and c->down == nullptr and c->right != nullptr)
		{
			Node<T>* lef = c->left;
			Node<T>* righ = c->right;
			while (lef != nullptr)
			{
				lef->right = righ;
				righ->left = lef;
				righ = righ->down;
				lef = lef->down;
			}
		}
		else if (c->down == nullptr)
		{
			c = nullptr;
		}
		else
		{
			Node<T>* temp = c;
			while (temp->down->down != nullptr)
			{
				temp->data = temp->down->data;
				temp = temp->down;
			}
			if (temp->down->left != nullptr and temp->down->right != nullptr)
			{
				temp->down->left->right = nullptr;
				temp->down->right->left = nullptr;
			}
			temp->down = nullptr;
		}
	}

	void deleteColumn()
	{
		Node<T>* temp = this->current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		if (this->current->left == nullptr and head->right != nullptr)
		{
			head = head->right;
			this->current = this->current->right;
		}
		if (this->current->left != nullptr and this->current->right != nullptr)
		{
			temp = temp->left;
			while (temp != nullptr)
			{
				temp->right->right->left = temp;
				temp->right = temp->right->right;
				temp = temp->down;
			}

			this->current = this->current->left;
		}
		else if (this->current->right == nullptr)
		{
			temp = temp->left;
			while (temp != nullptr)
			{
				temp->right = nullptr;
				temp = temp->down;
			}
			this->current = this->current->left;
		}
		else if (this->current->left == nullptr)
		{
			while (temp != nullptr)
			{
				temp->right->left = nullptr;
			}
			this->current = this->current->right;
		}
		this->columns--;
	}

	void deleteRowCurrent()
	{
		Node<T>* temp = this->current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		if (temp == head and head->down != nullptr)
		{
			this->current = this->current->down;
			head = head->down;
			temp = temp->down;
			while (temp != nullptr)
			{
				temp->up = nullptr;
				temp = temp->right;
			}
		}
		else if (this->current->down == nullptr)
		{
			temp = temp->up;
			this->current = this->current->up;
			while (temp != nullptr)
			{
				temp->down = nullptr;
				temp = temp->right;
			}
		}
		else
		{
			temp = temp->up;
			this->current = this->current->up;
			while (temp != nullptr)
			{
				temp->down->down->up = temp;
				temp->down = temp->down->down;
				temp = temp->right;
			}
		}
		this->rows--;
	}

	void ClearRow()
	{
		Node<T>* temp = this->current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		while (temp != nullptr)
		{
			temp->data = "0";
			temp = temp->right;
		}
	}

	void ClearColumn()
	{
		Node<T>* temp = current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		while (temp != nullptr)
		{
			temp->data = "0";
			temp = temp->down;
		}
	}

	// selected sum
	double GetRangeSum(rangeStart start, rangeEnd end)
	{
		double sum = 0;
		iiterator iterator(this->head);
		for (int i = 1; i < start.rowStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			iterator++;
		}

		for (int i = 1; i < start.colStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			++iterator;
		}

		for (int row = start.rowStart; row <= end.rowEnd; row++)
		{
			for (int col = start.colStart; col <= end.colEnd; col++)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				if (isStrDigit(iterator.iter->data))
				{
					sum += stod(iterator.iter->data);
				}
				iterator++;
			}
			for (int i = end.colEnd; i >= start.colStart; i--)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				iterator--;
			}
			++iterator;
		}

		return sum;
	}

	// selected average
	double GetRangeAverage(rangeStart start, rangeEnd end)
	{
		int cellCount = (end.rowEnd - start.rowStart + 1) * (end.colEnd - start.colStart + 1);
		double sum = GetRangeSum(start, end);
		double average = sum / cellCount;

		return average;
	}
	// selected Count
	int getSelectedCount(rangeStart start, rangeEnd end)
	{
		int cellCount = (end.rowEnd - start.rowStart + 1) * (end.colEnd - start.colStart + 1);
		return cellCount;
	}

	// Give Minimum value
	double GetRangeMinimum(rangeStart start, rangeEnd end)
	{
		double minimum = INT64_MAX;
		iiterator iterator(this->head);
		for (int i = 1; i < start.rowStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			iterator++;
		}

		for (int i = 1; i < start.colStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			++iterator;
		}

		for (int row = start.rowStart; row <= end.rowEnd; row++)
		{
			for (int col = start.colStart; col <= end.colEnd; col++)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				if(isStrDigit(iterator.iter->data))
				minimum = min(stod(iterator.iter->data), minimum);
				iterator++;
			}
			for (int i = end.colEnd; i >= start.colStart; i--)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				iterator--;
			}
			++iterator;
		}

		return minimum;
	}
	double GetRangeMaximum(rangeStart start, rangeEnd end)
	{
		double maximum = INT64_MIN;
		iiterator iterator(this->head);
		for (int i = 1; i < start.rowStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			iterator++;
		}

		for (int i = 1; i < start.colStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return 0.0;
			}
			++iterator;
		}

		for (int row = start.rowStart; row <= end.rowEnd; row++)
		{
			for (int col = start.colStart; col <= end.colEnd; col++)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				if (isStrDigit(iterator.iter->data))
				maximum = max(stod(iterator.iter->data), maximum);
				iterator++;
			}
			for (int i = end.colEnd; i >= start.colStart; i--)
			{
				if (iterator.iter == nullptr)
				{
					return 0.0;
				}
				iterator--;
			}
			++iterator;
		}

		return maximum;
	}

	// Give Maximum value
	void RangeCopySelect(rangeStart start, rangeEnd end)
	{
		// vector<string> v;
		iiterator iterator(this->head);
		for (int i = 1; i < start.rowStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return;
			}
			iterator++;
		}

		for (int i = 1; i < start.colStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return;
			}
			++iterator;
		}

		for (int row = start.rowStart; row <= end.rowEnd; row++)
		{
			for (int col = start.colStart; col <= end.colEnd; col++)
			{
				if (iterator.iter == nullptr)
				{
					return;
				}
				pasteVecter.push_back(iterator.iter->data);
				iterator++;
			}
			pasteVecter.push_back("$$$");
			for (int i = end.colEnd; i >= start.colStart; i--)
			{
				if (iterator.iter == nullptr)
				{
					return;
				}
				iterator--;
			}
			++iterator;
		}
		pasteVecter.pop_back();
	}
	void RangeCutSelect(rangeStart start, rangeEnd end)
	{
		// vector<string> v;
		iiterator iterator(this->head);
		for (int i = 1; i < start.rowStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return;
			}
			iterator++;
		}

		for (int i = 1; i < start.colStart; i++)
		{
			if (iterator.iter == nullptr)
			{
				return;
			}
			++iterator;
		}

		for (int row = start.rowStart; row <= end.rowEnd; row++)
		{
			for (int col = start.colStart; col <= end.colEnd; col++)
			{
				if (iterator.iter == nullptr)
				{
					return;
				}
				pasteVecter.push_back(iterator.iter->data);
				iterator.iter->data = "";
				iterator++;
			}
			pasteVecter.push_back("$$$");
			for (int i = end.colEnd; i >= start.colStart; i--)
			{
				if (iterator.iter == nullptr)
				{
					return;
				}
				iterator--;
			}
			++iterator;
		}
		pasteVecter.pop_back();
	}

	void paste()
	{
		if (this->pasteVecter.empty())
		{
			return;
		}
		else
		{
			Node<T>* temp = this->current;



			for (int i = 0; i < pasteVecter.size(); i++)
			{
				if (pasteVecter[i] == "$$$" && this->current->down != nullptr)
				{
					this->current = this->current->down;
					temp = this->current;
					temp->data = pasteVecter[i];
				}
				else if (pasteVecter[i] == "$$$" && this->current->down == nullptr)
				{
					InserRowBelow();
					this->current = this->current->down;
					temp = this->current;
				}
				else
				{
					temp->data = pasteVecter[i];

					if (i < pasteVecter.size() - 1 && pasteVecter[i + 1] != "$$$" && temp->right == nullptr)
					{
						insertLastColumn();
					}

					temp = temp->right;
				}
			}
		}
	}

	void saveToFile(string file)
	{
		fstream f(file, ios::out);

		if (!f.is_open())
		{
			cout << "Error opening the file named: " << file << endl;
			return;
		}
		else
		{
			Node<T>* Column = head;
			Node<T>* Row = head;

			while (Row != nullptr)
			{
				while (Column != nullptr)
				{
					f << Column->data;
					if (Column->right != nullptr)
						f << ",";
					Column = Column->right;
				}
				f << endl;
				Row = Row->down;
				Column = Row;
			}

			f.close();
		}
	}

	// template <typename T>
	bool loadExcelFromFile(excel<T>& myExcel, const string& file)
	{
		ifstream inputFile(file);
		if (!inputFile.is_open())
		{
			cerr << "Error opening the file named: " << file << endl;
			return false;
		}

		string line;
		int rowCount = 0;
		int colCount = 0;

		Node<T>* prevRowNode = nullptr;
		Node<T>* firstNodeInRow = nullptr;

		while (getline(inputFile, line))
		{
			istringstream iss(line);
			string cell;

			Node<T>* prevColNode = nullptr;

			while (getline(iss, cell, ','))
			{
				Node<T>* newNode = new Node<T>(cell);

				if (prevColNode != nullptr)
				{
					newNode->left = prevColNode;
					prevColNode->right = newNode;
				}

				if (prevRowNode != nullptr)
				{
					newNode->up = prevRowNode;
					prevRowNode->down = newNode;
				}

				if (firstNodeInRow == nullptr)
				{
					firstNodeInRow = newNode;
				}

				prevColNode = newNode;

				if (colCount > 0)
				{
					Node<T>* correspondingNode = prevColNode;
					while (correspondingNode->up != nullptr && newNode->up == nullptr)
					{
						correspondingNode = correspondingNode->up;
					}
					newNode->up = correspondingNode;
					correspondingNode->down = newNode;
				}
			}

			if (rowCount > 0 && firstNodeInRow != nullptr)
			{
				Node<T>* correspondingNode = myExcel.head;
				for (int i = 0; i < colCount; ++i)
				{
					correspondingNode = correspondingNode->down;
				}
				while (correspondingNode->down != nullptr)
				{
					correspondingNode = correspondingNode->down;
					correspondingNode->up = firstNodeInRow;
					firstNodeInRow->down = correspondingNode;
					firstNodeInRow = firstNodeInRow->down;
				}

				firstNodeInRow = nullptr;
			}

			++rowCount;
			colCount = prevRowNode == nullptr ? colCount : prevRowNode->data.empty() ? colCount
				: colCount + 1;

			prevRowNode = firstNodeInRow;
		}

		myExcel.rows = rowCount;
		myExcel.columns = colCount;

		inputFile.close();
		return true;
	}

private:
	// Helper function to clear existing data
	void clearExcel()
	{
		this->ex = nullptr;
		this->head = nullptr;
		this->current = nullptr;
		this->rows = 0;
		this->columns = 0;
		this->pasteVecter.clear();
	}

	// Helper function to get the node at a specific position
	Node<T>* getNodeAtPosition(int row, int col)
	{
		Node<T>* temp = this->head;
		for (int i = 1; i < row; i++)
		{
			temp = temp->down;
		}
		for (int i = 1; i < col; i++)
		{
			temp = temp->right;
		}
		return temp;
	}

	bool isStrDigit(string& str)
	{
		for (auto c : str)
		{
			if (isdigit(c))
			{
				continue;
			}
			else {
				return false;
			}
		}
		return true;
	}
};
