#pragma once
#include <iostream>
#include <iomanip>
#include <string>
#include <time.h>
#include <sstream>

using namespace std;

namespace FormTest {
	using namespace Excel;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MyForm の概要
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: ここにコンストラクター コードを追加します
			//
		}

	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::Button^  button2;
	protected:
	private: System::Windows::Forms::DataGridView^  dataGridView1;
	private: System::Windows::Forms::TextBox^  textBox1;
	private: System::Windows::Forms::TextBox^  textBox2;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::Label^  label2;

	private:
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			this->textBox2 = (gcnew System::Windows::Forms::TextBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(458, 166);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(99, 28);
			this->button1->TabIndex = 0;
			this->button1->Text = L"ADD";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			//
			// button2
			// 
			this->button2->Location = System::Drawing::Point(590, 166);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(99, 28);
			this->button2->TabIndex = 0;
			this->button2->Text = L"Delete";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Location = System::Drawing::Point(30, 41);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->RowTemplate->Height = 21;
			this->dataGridView1->Size = System::Drawing::Size(395, 460);
			this->dataGridView1->TabIndex = 1;
			this->dataGridView1->ColumnCount = 3;
			this->dataGridView1->Columns[0]->Name = "Date";
			this->dataGridView1->Columns[1]->Name = "ID";
			this->dataGridView1->Columns[2]->Name = "Name";
			this->dataGridView1->AutoSizeRowsMode = DataGridViewAutoSizeRowsMode::DisplayedCellsExceptHeaders;
			this->dataGridView1->SelectionMode = DataGridViewSelectionMode::FullRowSelect;
			// 
			// textBox1
			// 
			this->textBox1->Location = System::Drawing::Point(458, 63);
			this->textBox1->Name = L"textBox1";
			this->textBox1->Size = System::Drawing::Size(57, 19);
			this->textBox1->TabIndex = 2;
			// 
			// textBox2
			// 
			this->textBox2->Location = System::Drawing::Point(458, 111);
			this->textBox2->Name = L"textBox2";
			this->textBox2->Size = System::Drawing::Size(125, 19);
			this->textBox2->TabIndex = 3;
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(456, 45);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(59, 12);
			this->label1->TabIndex = 4;
			this->label1->Text = L"ID Number";
			this->label1->Click += gcnew System::EventHandler(this, &MyForm::label1_Click);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(456, 94);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(34, 12);
			this->label2->TabIndex = 5;
			this->label2->Text = L"Name";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(889, 539);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->textBox2);
			this->Controls->Add(this->textBox1);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->button2);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void label1_Click(System::Object^  sender, System::EventArgs^  e) {
	}
private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {

	time_t timer = time(0);
	struct tm *timeStruct = localtime(&timer);

	int year = 1900 + timeStruct->tm_year;
	int month = 1 + timeStruct->tm_mon;
	int day = timeStruct->tm_mday;
	int hour = timeStruct->tm_hour;
	int minute = timeStruct->tm_min;
	int seconds = timeStruct->tm_sec;
	String^ idNum = textBox1->Text;
	String^ name = textBox2->Text;

	stringstream ss;
	ss << year << "/" << setw(2) << setfill('0') << month << "/" << setw(2) << setfill('0') << day << "," << setw(2) << setfill('0') << hour << ":" << setw(2) << setfill('0') << minute << ":" << setw(2) << setfill('0') << seconds;
	string getTime = ss.str();
	String^ dateTime = gcnew String(getTime.c_str());

	array<String^>^ rows0 = gcnew array<String^>{dateTime,idNum,name};
	DataGridViewRowCollection^ rows = this->dataGridView1->Rows;
	rows->Add(rows0);
}

	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {
	}

};
}
