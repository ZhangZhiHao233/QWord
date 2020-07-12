#ifndef QWORD_H
#define QWORD_H

#include <QObject>
#include <QAxObject>
#include <QList>

struct openedDocs
{
	QAxObject *pDoc;
	QString docName;
	int iRows;
};

class QWord : public QObject
{
	Q_OBJECT

public:
	QWord(QObject *parent = 0);
	~QWord();

public:
	bool InitWord();
	void closeWord();//œ»CloseDoc‘⁄closeWord

	QAxObject* CreateDoc(const QString& docName);
	QAxObject* OpenDoc(const QString& docName);
	void CloseDoc(QAxObject* &pDoc);
	bool DelDoc(const QString& docName);

	void InsertTitle(QAxObject* &pDoc, const QString& title);
	void InsertInfo(QAxObject* &pDoc, const QString& info);
	
private:
	QAxObject* IsOpened(const QString& docName);
	void AddList(QAxObject* &pDoc, const QString& docName);
	bool DelList(QAxObject* &pDoc);
	QString GetNameByObject(QAxObject* &pDoc);
	int GetUsedRows(QAxObject* &pDoc);
	void SetUsedRows(QAxObject* &pDoc, int iRows);

private:
	QAxObject *pWord;
	QAxObject *pDocs;
	QList<openedDocs> openedDocList;
};

#endif // QWORD_H
