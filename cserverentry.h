#ifndef CSERVERENTRY_H
#define CSERVERENTRY_H


#include <QString>
#include <QDate>

#include <QList>

#include <QMetaType>


class cServerEntry
{
public:
	cServerEntry();

	bool		load(const QString& fileName);

//private:
	QString		m_fileName;
	QString		m_serverName;
	QDate		m_reportDate;
	QString		m_customerInterface;
	qint16		m_totalNumberOfUsers;
	qint16		m_userDisabled;
	qreal		m_availableDiskSpace;
	qreal		m_usedDiskSpace;
	qreal		m_usedDiskSpacePercent;
	qreal		m_maximumDiskUtil;
	qreal		m_growthBasePrevMonth;
	qreal		m_growthBaseStartDate;
	qreal		m_mainStorage;
	qint16		m_numberOfProblemRecords;
	qreal		m_availability;
	qreal		m_averageCPUUtil;
	QString		m_averageCPUUtilTimeframe;
	qreal		m_maximumCPUUtil;
	qint32		m_downtimeCum;

};

Q_DECLARE_METATYPE(cServerEntry*)

class cServerEntryList : public QList<cServerEntry*>
{
public:
	explicit				cServerEntryList();

	bool					loadPath(const QString& path);
	bool					save(const QString& fileName);

signals:

public slots:
};

#endif // CSERVERENTRY_H
