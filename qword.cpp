#include "qword.h"
#include <QFile>
#include <QVariant>
#include <QDir>
#include <QDebug>
#include <QByteArray>

QWord::QWord(QObject *parent)
	: QObject(parent)
{
	
}

QWord::~QWord()
{
}

bool QWord::InitWord()
{
	pWord = new QAxObject("Word.Application");
	pWord->dynamicCall("SetVisible(bool)", "true");
	pWord->setProperty("DisplayAlerts", false);
	pDocs = pWord->querySubObject("Documents");
	return true;
}

void QWord::closeWord()
{
	pWord->dynamicCall("Quit()");
}

QAxObject* QWord::CreateDoc(const QString& docName)
{
	QString filePath, docFilePath, dotFilePath;
	QDir fileDir;
	filePath = QDir::currentPath();
	filePath.replace("/", "\\");
	filePath.append("\\");

	dotFilePath = filePath + QString("myDot.dotx");
	docFilePath = filePath + docName;

	QFile fileDot(dotFilePath);
	QFile fileDoc(docFilePath);

	QAxObject *pDoc = NULL;
	if (!fileDot.exists())
	{
		qDebug()<<"No Dot file!";
		return pDoc;
	}

	if(!fileDoc.exists())
	{
		pDocs->dynamicCall("Add (QString)", dotFilePath);
		pDoc = pWord->querySubObject("ActiveDocument");
		pDoc->dynamicCall("SaveAs (const QString&)", QDir::toNativeSeparators(docFilePath));
		AddList(pDoc, docName);
	}
	else
	{
		pDocs->dynamicCall("Open (const QString&)", docFilePath);
		pDoc = pWord->querySubObject("ActiveDocument");
		AddList(pDoc, docName);

	}

	return pDoc;
}

QAxObject* QWord::OpenDoc(const QString& docName)
{
	QString filePath, docFilePath, dotFilePath;
	QDir fileDir;
	filePath = QDir::currentPath();
	filePath.replace("/", "\\");
	filePath.append("\\");

	docFilePath = filePath + docName;
	QFile fileDoc(docFilePath);
	QAxObject* pDoc;
	if(!fileDoc.exists())
	{
		pDoc = NULL;
	}
	else
	{
		pDoc = IsOpened(docName);
		if (!pDoc)
		{
			pDocs->dynamicCall("Open (const QString&)", docFilePath);
			pDoc = pDocs->querySubObject("ActiveDocument");
			AddList(pDoc, docName);
		}
	}

	return pDoc;
}

void QWord::CloseDoc(QAxObject* &pDoc)
{
	pDoc->dynamicCall("Save()");
	pDoc->dynamicCall("Close()");
	DelList(pDoc);
}

bool QWord::DelDoc(const QString& docName)
{
	QString filePath, docFilePath, dotFilePath;
	QDir fileDir;
	filePath = QDir::currentPath();
	filePath.replace("/", "\\");
	filePath.append("\\");

	docFilePath = filePath + docName;
	QFile fileDoc(docFilePath);
	if(fileDoc.exists())
	{
		fileDoc.remove();
		return true;
	}
	else
	{
		return false;
	}
}

void QWord::InsertTitle(QAxObject* &pDoc, const QString& title)
{
	QString cmd;
	char *temp = nullptr;
	int iRows = GetUsedRows(pDoc);
	if (iRows != 1)
	{
		iRows += 1;
	}
	else
	{
		iRows = 1;
	}

	cmd = QString("Bookmarks(label_%1)").arg(iRows);
	QByteArray ba = cmd.toLatin1();
	temp = ba.data();
	QAxObject* label = pDoc->querySubObject(temp);

	if (!label->isNull())
	{
		label->dynamicCall("Select(void)");

		QAxObject* selection = pWord->querySubObject("Selection");
		selection->querySubObject("Font")->setProperty("Color", 255);
		selection->querySubObject("Font")->setProperty("Name", QString("ËÎÌו"));
		selection->querySubObject("Font")->setProperty("Size", 20);
		selection->querySubObject("Paragraphs(int)", 1)->setProperty("Alignment", QString("wdAlignParagraphCenter"));
		label->querySubObject("Range")->setProperty("Text", title);
		pDoc->dynamicCall("Save()");
		iRows++;
		SetUsedRows(pDoc, iRows);
	}
}

void QWord::InsertInfo(QAxObject* &pDoc, const QString& info)
{
	QString cmd;
	char *temp = nullptr;
	int iRows = GetUsedRows(pDoc);
	cmd = QString("Bookmarks(label_%1)").arg(iRows);
	QByteArray ba = cmd.toLatin1();
	temp = ba.data();
	QAxObject* label = pDoc->querySubObject(temp);
	if (!label->isNull())
	{
		label->dynamicCall("Select(void)");

		QAxObject* selection = pWord->querySubObject("Selection");
		selection->querySubObject("Font")->setProperty("Color", 0);
		selection->querySubObject("Font")->setProperty("Name", QString("ËÎÌו"));
		selection->querySubObject("Font")->setProperty("Size", 14);
		selection->querySubObject("Paragraphs(int)", 1)->setProperty("Alignment", QString("wdAlignParagraphLeft"));
		label->querySubObject("Range")->setProperty("Text", info);
		pDoc->dynamicCall("Save()");
		iRows++;
		SetUsedRows(pDoc, iRows);
	}
}

QAxObject* QWord::IsOpened(const QString& docName)
{
	for (int i = 0; i< openedDocList.size(); i++)
	{
		if (openedDocList.at(i).docName == docName)
		{
			return openedDocList.at(i).pDoc;
		}
	}

	return NULL;

}

void QWord::AddList(QAxObject* &pDoc, const QString& docName)
{
	openedDocs w;
	w.pDoc = pDoc;
	w.docName = docName;
	w.iRows = 1;
	openedDocList.append(w);
}

bool QWord::DelList(QAxObject* &pDoc)
{
	int iSize = openedDocList.size();
	for (int i = 0; i< openedDocList.size(); i++)
	{
		if (openedDocList.at(i).pDoc == pDoc)
		{
			if (iSize > 1)
			{
				delete pDoc;
			}

			pDoc = NULL;
			openedDocList.removeAt(i);
			return true;
		}
	}
	return false;
}

QString QWord::GetNameByObject(QAxObject* &pDoc)
{
	if (pDoc == NULL)
	{
		return NULL;
	}

	for (int i = 0; i< openedDocList.size(); i++)
	{
		if (openedDocList.at(i).pDoc == pDoc)
		{
			return openedDocList.at(i).docName;
		}
	}

	return NULL;
}

int QWord::GetUsedRows(QAxObject* &pDoc)
{
	for (int i = 0; i< openedDocList.size(); i++)
	{
		if (openedDocList.at(i).pDoc == pDoc)
		{
			return openedDocList.at(i).iRows;
		}
	}

	return -1;
}

void QWord::SetUsedRows(QAxObject* &pDoc, int iRows)
{
	QList<openedDocs>::iterator itr;
	for(itr =openedDocList.begin(); itr!=openedDocList.end(); ++itr)
	{
		if (itr->pDoc == pDoc)
		{
			itr->iRows = iRows;
		}
	}
}
